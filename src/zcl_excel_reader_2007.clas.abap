class zcl_excel_reader_2007 definition
  public
  create public .

  public section.
*"* public components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!

    interfaces zif_excel_reader .

    class-methods fill_struct_from_attributes
      importing
        !ip_element   type ref to if_ixml_element
      changing
        !cp_structure type any .
  protected section.

    types:
*"* protected components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!
      begin of t_relationship,
        id           type string,
        type         type string,
        target       type string,
        targetmode   type string,
        worksheet    type ref to zcl_excel_worksheet,
        sheetid      type string,     "ins #235 - repeat rows/cols - needed to identify correct sheet
        localsheetid type string,
      end of t_relationship .
    types:
      begin of t_fileversion,
        appname      type string,
        lastedited   type string,
        lowestedited type string,
        rupbuild     type string,
        codename     type string,
      end of t_fileversion .
    types:
      begin of t_sheet,
        name    type string,
        sheetid type string,
        id      type string,
        state   type string,
      end of t_sheet .
    types:
      begin of t_workbookpr,
        codename            type string,
        defaultthemeversion type string,
      end of t_workbookpr .
    types:
      begin of t_sheetpr,
        codename type string,
      end of t_sheetpr .
    types:
      begin of t_range,
        name         type string,
        hidden       type string,       "inserted with issue #235 because Autofilters didn't passthrough
        localsheetid type string,       " issue #163
      end of t_range .
    types:
      t_fills   type standard table of ref to zcl_excel_style_fill with non-unique default key .
    types:
      t_borders type standard table of ref to zcl_excel_style_borders with non-unique default key .
    types:
      t_fonts   type standard table of ref to zcl_excel_style_font with non-unique default key .
    types:
      t_style_refs type standard table of ref to zcl_excel_style with non-unique default key .
    types:
      begin of t_color,
        indexed type string,
        rgb     type string,
        theme   type string,
        tint    type string,
      end of t_color .
    types:
      begin of t_rel_drawing,
        id          type string,
        content     type xstring,
        file_ext    type string,
        content_xml type ref to if_ixml_document,
      end of t_rel_drawing .
    types:
      t_rel_drawings type standard table of t_rel_drawing with non-unique default key .
    types:
      begin of gts_external_hyperlink,
        id     type string,
        target type string,
      end of gts_external_hyperlink .
    types:
      gtt_external_hyperlinks type hashed table of gts_external_hyperlink with unique key id .
    types:
      begin of ty_ref_formulae,
        sheet   type ref to zcl_excel_worksheet,
        row     type i,
        column  type i,
        si      type i,
        ref     type string,
        formula type string,
      end   of ty_ref_formulae .
    types:
      tyt_ref_formulae type hashed table of ty_ref_formulae with unique key sheet row column .
    types:
      begin of t_shared_string,
        value type string,
        rtf   type zif_excel_data_decl=>zexcel_t_rtf,
      end of t_shared_string .
    types:
      t_shared_strings type standard table of t_shared_string with default key .
    types:
      begin of t_table,
        id     type string,
        target type string,
      end of t_table .
    types:
      t_tables type hashed table of t_table with unique key id .

    data shared_strings type t_shared_strings .
    data styles type t_style_refs .
    data mt_ref_formulae type tyt_ref_formulae .
    data mt_dxf_styles type zif_excel_data_decl=>zexcel_t_styles_cond_mapping .

    methods fill_row_outlines
      importing
        !io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods get_from_zip_archive
      importing
        !i_filename      type string
      returning
        value(r_content) type xstring
      raising
        zcx_excel .
    methods get_ixml_from_zip_archive
      importing
        !i_filename     type string
        !is_normalizing type abap_bool default 'X'
      returning
        value(r_ixml)   type ref to if_ixml_document
      raising
        zcx_excel .
    methods load_drawing_anchor
      importing
        !io_anchor_element   type ref to if_ixml_element
        !io_worksheet        type ref to zcl_excel_worksheet
        !it_related_drawings type t_rel_drawings .
    methods load_shared_strings
      importing
        !ip_path type string
      raising
        zcx_excel .
    methods load_styles
      importing
        !ip_path  type string
        !ip_excel type ref to zcl_excel
      raising
        zcx_excel .
    methods load_dxf_styles
      importing
        !iv_path  type string
        !io_excel type ref to zcl_excel
      raising
        zcx_excel .
    methods load_style_borders
      importing
        !ip_xml           type ref to if_ixml_document
      returning
        value(ep_borders) type t_borders .
    methods load_style_fills
      importing
        !ip_xml         type ref to if_ixml_document
      returning
        value(ep_fills) type t_fills .
    methods load_style_font
      importing
        !io_xml_element type ref to if_ixml_element
      returning
        value(ro_font)  type ref to zcl_excel_style_font .
    methods load_style_fonts
      importing
        !ip_xml         type ref to if_ixml_document
      returning
        value(ep_fonts) type t_fonts .
    methods load_style_num_formats
      importing
        !ip_xml               type ref to if_ixml_document
      returning
        value(ep_num_formats) type zcl_excel_style_number_format=>t_num_formats .
    methods load_workbook
      importing
        !iv_workbook_full_filename type string
        !io_excel                  type ref to zcl_excel
      raising
        zcx_excel .
    methods load_worksheet
      importing
        !ip_path      type string
        !io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_worksheet_cond_format
      importing
        !io_ixml_worksheet type ref to if_ixml_document
        !io_worksheet      type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_worksheet_cond_format_aa
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond.
    methods load_worksheet_cond_format_ci
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_cond_format_cs
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_cond_format_ex
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_cond_format_is
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_cond_format_db
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_cond_format_t10
      importing
        !io_ixml_rule  type ref to if_ixml_element
        !io_style_cond type ref to zcl_excel_style_cond .
    methods load_worksheet_drawing
      importing
        !ip_path      type string
        !io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_comments
      importing
        ip_path      type string
        io_worksheet type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_worksheet_hyperlinks
      importing
        !io_ixml_worksheet      type ref to if_ixml_document
        !io_worksheet           type ref to zcl_excel_worksheet
        !it_external_hyperlinks type gtt_external_hyperlinks
      raising
        zcx_excel .
    methods load_worksheet_ignored_errors
      importing
        !io_ixml_worksheet type ref to if_ixml_document
        !io_worksheet      type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_worksheet_pagebreaks
      importing
        !io_ixml_worksheet type ref to if_ixml_document
        !io_worksheet      type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    methods load_worksheet_autofilter
      importing
        io_ixml_worksheet type ref to if_ixml_document
        io_worksheet      type ref to zcl_excel_worksheet
      raising
        zcx_excel.
    methods load_worksheet_pagemargins
      importing
        !io_ixml_worksheet type ref to if_ixml_document
        !io_worksheet      type ref to zcl_excel_worksheet
      raising
        zcx_excel .
    "! <p class="shorttext synchronized" lang="en">Load worksheet tables</p>
    methods load_worksheet_tables
      importing
        io_ixml_worksheet type ref to if_ixml_document
        io_worksheet      type ref to zcl_excel_worksheet
        iv_dirname        type string
        it_tables         type t_tables
      raising
        zcx_excel .
    class-methods resolve_path
      importing
        !ip_path         type string
      returning
        value(rp_result) type string .
    methods resolve_referenced_formulae .
    methods unescape_string_value
      importing
        i_value       type string
      returning
        value(result) type string.
    methods get_dxf_style_guid
      importing
        !io_ixml_dxf         type ref to if_ixml_element
        !io_excel            type ref to zcl_excel
      returning
        value(rv_style_guid) type zif_excel_data_decl=>zexcel_cell_style .
    methods load_theme
      importing
        iv_path   type string
        !ip_excel type ref to zcl_excel
      raising
        zcx_excel.
    methods provided_string_is_escaped
      importing
        !value            type string
      returning
        value(is_escaped) type abap_bool.

    constants: begin of namespace,
                 x14ac            type string value 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
                 vba_project      type string value 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', "#EC NEEDED     for future incorporation of XLSM-reader
                 c                type string value 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                 a                type string value 'http://schemas.openxmlformats.org/drawingml/2006/main',
                 xdr              type string value 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                 mc               type string value 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                 r                type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                 chart            type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                 drawing          type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                 hyperlink        type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                 image            type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                 office_document  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                 printer_settings type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
                 shared_strings   type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                 styles           type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                 theme            type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                 worksheet        type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                 relationships    type string value 'http://schemas.openxmlformats.org/package/2006/relationships',
                 core_properties  type string value 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                 main             type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               end of namespace.

  private section.

    data zip type ref to lcl_zip_archive .
    data: gid type i.

    methods create_zip_archive
      importing
        !i_xlsx_binary       type xstring
        !i_use_alternate_zip type clike optional
      returning
        value(e_zip)         type ref to lcl_zip_archive
      raising
        zcx_excel .
    methods read_from_applserver
      importing
        !i_filename         type csequence
      returning
        value(r_excel_data) type xstring
      raising
        zcx_excel.
    methods read_from_local_file
      importing
        !i_filename         type csequence
      returning
        value(r_excel_data) type xstring
      raising
        zcx_excel .
endclass.



class zcl_excel_reader_2007 implementation.


  method create_zip_archive.
    case i_use_alternate_zip.
      when space.
        e_zip = lcl_abap_zip_archive=>create( i_xlsx_binary ).
      when others.
        e_zip = lcl_alternate_zip_archive=>create( i_data                = i_xlsx_binary
                                                   i_alternate_zip_class = i_use_alternate_zip ).
    endcase.
  endmethod.


  method fill_row_outlines.

    types: begin of lts_row_data,
             row           type i,
             outline_level type i,
           end of lts_row_data,
           ltt_row_data type sorted table of lts_row_data with unique key row.

    data: lt_row_data             type ltt_row_data,
          ls_row_data             like line of lt_row_data,
          lt_collapse_rows        type hashed table of i with unique key table_line,
          lv_collapsed            type abap_bool,
          lv_outline_level        type i,
          lv_next_consecutive_row type i,
          lt_outline_rows         type zcl_excel_worksheet=>mty_ts_outlines_row,
          ls_outline_row          like line of lt_outline_rows,
          lo_row                  type ref to zcl_excel_row,
          lo_row_iterator         type ref to zcl_excel_collection_iterator,
          lv_row_offset           type i,
          lv_row_collapse_flag    type i.


    field-symbols: <ls_row_data>      like line of lt_row_data.

* First collect information about outlines ( outline leven and collapsed state )
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    while lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      ls_row_data-row           = lo_row->get_row_index( ).
      ls_row_data-outline_level = lo_row->get_outline_level( ).
      if ls_row_data-outline_level is not initial.
        insert ls_row_data into table lt_row_data.
      endif.

      lv_collapsed = lo_row->get_collapsed( ).
      if lv_collapsed = abap_true.
        insert lo_row->get_row_index( ) into table lt_collapse_rows.
      endif.
    endwhile.

* Now parse this information - we need consecutive rows - any gap will create a new outline
    do 7 times.  " max number of outlines allowed
      lv_outline_level = sy-index.
      clear lv_next_consecutive_row.
      clear ls_outline_row.
      loop at lt_row_data assigning <ls_row_data> where outline_level >= lv_outline_level. "#EC CI_SORTSEQ

        if lv_next_consecutive_row    <> <ls_row_data>-row   " A gap --> close all open outlines
          and lv_next_consecutive_row is not initial.        " First time in loop.
          insert ls_outline_row into table lt_outline_rows.
          clear: ls_outline_row.
        endif.

        if ls_outline_row-row_from is initial.
          ls_outline_row-row_from = <ls_row_data>-row.
        endif.
        ls_outline_row-row_to = <ls_row_data>-row.

        lv_next_consecutive_row = <ls_row_data>-row + 1.

      endloop.
      if ls_outline_row-row_from is not initial.
        insert ls_outline_row into table lt_outline_rows.
      endif.
    enddo.

* lt_outline_rows holds all outline information
* we now need to determine whether the outline is collapsed or not
    loop at lt_outline_rows into ls_outline_row.

      if io_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off.
        lv_row_collapse_flag = ls_outline_row-row_from - 1.
      else.
        lv_row_collapse_flag = ls_outline_row-row_to + 1.
      endif.
      read table lt_collapse_rows transporting no fields with table key table_line = lv_row_collapse_flag.
      if sy-subrc = 0.
        ls_outline_row-collapsed = abap_true.
      endif.
      io_worksheet->set_row_outline( iv_row_from  = ls_outline_row-row_from
                                     iv_row_to    = ls_outline_row-row_to
                                     iv_collapsed = ls_outline_row-collapsed ).

    endloop.

* Finally purge outline information ( collapsed state, outline leve)  from row_dimensions, since we want to keep these in the outline-table
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    while lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).

      lo_row->set_outline_level( 0 ).
      lo_row->set_collapsed( abap_false ).

    endwhile.

  endmethod.


  method fill_struct_from_attributes.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

    data: lv_name       type string,
          lo_attributes type ref to if_ixml_named_node_map,
          lo_attribute  type ref to if_ixml_attribute,
          lo_iterator   type ref to if_ixml_node_iterator.

    field-symbols: <component>                  type any.

*--------------------------------------------------------------------*
* The values of named attributes of a tag are being read and moved into corresponding
* fields of given structure
* Behaves like move-corresonding tag to structure

* Example:
*     <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*   Here the attributes are Target, Type and Id.  Thus if the passed
*   structure has fieldnames Id and Target these would be filled with
*   "rId3" and "docProps/app.xml" respectively
*--------------------------------------------------------------------*
    clear cp_structure.

    lo_attributes  = ip_element->get_attributes( ).
    lo_iterator    = lo_attributes->create_iterator( ).
    lo_attribute  ?= lo_iterator->get_next( ).
    while lo_attribute is bound.

      lv_name = lo_attribute->get_name( ).
      translate lv_name to upper case.
      assign component lv_name of structure cp_structure to <component>.
      if sy-subrc = 0.
        <component> = lo_attribute->get_value( ).
      endif.
      lo_attribute ?= lo_iterator->get_next( ).

    endwhile.


  endmethod.


  method get_dxf_style_guid.
    data: lo_ixml_dxf_children          type ref to if_ixml_node_list,
          lo_ixml_iterator_dxf_children type ref to if_ixml_node_iterator,
          lo_ixml_dxf_child             type ref to if_ixml_element,

          lv_dxf_child_type             type string,

          lo_ixml_element               type ref to if_ixml_element,
          lo_ixml_element2              type ref to if_ixml_element,
          lv_val                        type string.

    data: ls_cstyle  type zif_excel_data_decl=>zexcel_s_cstyle_complete,
          ls_cstylex type zif_excel_data_decl=>zexcel_s_cstylex_complete.



    lo_ixml_dxf_children = io_ixml_dxf->get_children( ).
    lo_ixml_iterator_dxf_children = lo_ixml_dxf_children->create_iterator( ).
    lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).
    while lo_ixml_dxf_child is bound.

      lv_dxf_child_type = lo_ixml_dxf_child->get_name( ).
      case lv_dxf_child_type.

        when 'font'.
*--------------------------------------------------------------------*
* italic
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'i' uri = namespace-main ).
          if lo_ixml_element is bound.
            clear lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            if lv_val <> '0'.
              ls_cstyle-font-italic  = 'X'.
              ls_cstylex-font-italic = 'X'.
            endif.

          endif.
*--------------------------------------------------------------------*
* bold
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'b' uri = namespace-main ).
          if lo_ixml_element is bound.
            clear lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            if lv_val <> '0'.
              ls_cstyle-font-bold  = 'X'.
              ls_cstylex-font-bold = 'X'.
            endif.

          endif.
*--------------------------------------------------------------------*
* strikethrough
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'strike' uri = namespace-main ).
          if lo_ixml_element is bound.
            clear lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            if lv_val <> '0'.
              ls_cstyle-font-strikethrough  = 'X'.
              ls_cstylex-font-strikethrough = 'X'.
            endif.

          endif.
*--------------------------------------------------------------------*
* color
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'color' uri = namespace-main ).
          if lo_ixml_element is bound.
            clear lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'rgb' ).
            ls_cstyle-font-color-rgb  = lv_val.
            ls_cstylex-font-color-rgb = 'X'.
          endif.

        when 'fill'.
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'patternFill' uri = namespace-main ).
          if lo_ixml_element is bound.
            lo_ixml_element2 = lo_ixml_dxf_child->find_from_name_ns( name = 'bgColor' uri = namespace-main ).
            if lo_ixml_element2 is bound.
              clear lv_val.
              lv_val  = lo_ixml_element2->get_attribute_ns( 'rgb' ).
              if lv_val is not initial.
                ls_cstyle-fill-filltype       = zcl_excel_style_fill=>c_fill_solid.
                ls_cstyle-fill-bgcolor-rgb    = lv_val.
                ls_cstylex-fill-filltype      = 'X'.
                ls_cstylex-fill-bgcolor-rgb   = 'X'.
              endif.
              clear lv_val.
              lv_val  = lo_ixml_element2->get_attribute_ns( 'theme' ).
              if lv_val is not initial.
                ls_cstyle-fill-filltype         = zcl_excel_style_fill=>c_fill_solid.
                ls_cstyle-fill-bgcolor-theme    = lv_val.
                ls_cstylex-fill-filltype        = 'X'.
                ls_cstylex-fill-bgcolor-theme   = 'X'.
              endif.
              clear lv_val.
            endif.
          endif.

      endcase.

      lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).

    endwhile.

    rv_style_guid = io_excel->get_static_cellstyle_guid( ip_cstyle_complete  = ls_cstyle
                                                         ip_cstylex_complete = ls_cstylex  ).


  endmethod.


  method get_from_zip_archive.

    assert zip is bound. " zip object has to exist at this point

    r_content = zip->read(  i_filename ).

  endmethod.


  method get_ixml_from_zip_archive.
* The corresponding part of SAP note 2922674:
* A workaround for the replacement of characters of the supplementary plane in SAP_BASIS 7.51 or lower is to convert
* the UTF-8 value to ABAP variable type STRING using method CL_ABAP_CODEPAGE=>CONVERT_FROM and then to parse the
* document using an iXML input stream created with factory method IF_IXML_STREAM_FACTORY=>CREATE_ISTREAM_STRING.
* Do not use method IF_IXML_STREAM_FACTORY=>CREATE_ISTREAM_CSTRING in this context as it shows the unwanted behaviour.

    data: lv_content       type xstring,
          lo_ixml          type ref to if_ixml_core,
          lo_streamfactory type ref to if_ixml_stream_factory_core,
          lo_istream       type ref to if_ixml_istream_core,
          lo_parser        type ref to if_ixml_parser_core.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*
    lv_content        = me->get_from_zip_archive( i_filename ).
    lo_ixml           = cl_ixml_core=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    r_ixml            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = r_ixml ).
    lo_parser->set_normalizing( is_normalizing ).
    lo_parser->set_validating( mode = if_ixml_parser_core=>co_no_validation ).
    lo_parser->parse( ).

  endmethod.


  method load_drawing_anchor.

    types: begin of t_c_nv_pr,
             name type string,
             id   type string,
           end of t_c_nv_pr.

    types: begin of t_blip,
             cstate type string,
             embed  type string,
           end of t_blip.

    types: begin of t_chart,
             id type string,
           end of t_chart.

    types: begin of t_ext,
             cx type string,
             cy type string,
           end of t_ext.

    constants: lc_xml_attr_true     type string value 'true',
               lc_xml_attr_true_int type string value '1'.
    constants: lc_rel_chart type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
               lc_rel_image type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.

    data: lo_drawing     type ref to zcl_excel_drawing,
          node           type ref to if_ixml_element,
          node2          type ref to if_ixml_element,
          node3          type ref to if_ixml_element,
          node4          type ref to if_ixml_element,

          ls_upper       type zif_excel_data_decl=>zexcel_drawing_location,
          ls_lower       type zif_excel_data_decl=>zexcel_drawing_location,
          ls_size        type zif_excel_data_decl=>zexcel_drawing_size,
          ext            type t_ext,
          lv_content     type xstring,
          lv_relation_id type string,
          lv_title       type string,

          cnvpr          type t_c_nv_pr,
          blip           type t_blip,
          chart          type t_chart,
          drawing_type   type zif_excel_data_decl=>zexcel_drawing_type,

          rel_drawing    type t_rel_drawing.

    node ?= io_anchor_element->find_from_name_ns( name = 'from' uri = namespace-xdr ).
    check node is not initial.
    node2 ?= node->find_from_name_ns( name = 'col' uri = namespace-xdr ).
    ls_upper-col = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'row' uri = namespace-xdr ).
    ls_upper-row = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'colOff' uri = namespace-xdr ).
    ls_upper-col_offset = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'rowOff' uri = namespace-xdr ).
    ls_upper-row_offset = node2->get_value( ).

    node ?= io_anchor_element->find_from_name_ns( name = 'ext' uri = namespace-xdr ).
    if node is initial.
      clear ls_size.
    else.
      me->fill_struct_from_attributes( exporting ip_element = node changing cp_structure = ext ).
      ls_size-width = ext-cx.
      ls_size-height = ext-cy.
      try.
          ls_size-width  = zcl_excel_drawing=>emu2pixel( ls_size-width ).
        catch cx_root.
      endtry.
      try.
          ls_size-height = zcl_excel_drawing=>emu2pixel( ls_size-height ).
        catch cx_root.
      endtry.
    endif.

    node ?= io_anchor_element->find_from_name_ns( name = 'to' uri = namespace-xdr ).
    if node is initial.
      clear ls_lower.
    else.
      node2 ?= node->find_from_name_ns( name = 'col' uri = namespace-xdr ).
      ls_lower-col = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'row' uri = namespace-xdr ).
      ls_lower-row = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'colOff' uri = namespace-xdr ).
      ls_lower-col_offset = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'rowOff' uri = namespace-xdr ).
      ls_lower-row_offset = node2->get_value( ).
    endif.

    node ?= io_anchor_element->find_from_name_ns( name = 'pic' uri = namespace-xdr ).
    if node is not initial.
      node2 ?= node->find_from_name_ns( name = 'nvPicPr' uri = namespace-xdr ).
      check node2 is not initial.
      node3 ?= node2->find_from_name_ns( name = 'cNvPr' uri = namespace-xdr ).
      check node3 is not initial.
      me->fill_struct_from_attributes( exporting ip_element = node3 changing cp_structure = cnvpr ).
      lv_title = cnvpr-name.

      node2 ?= node->find_from_name_ns( name = 'blipFill' uri = namespace-xdr ).
      check node2 is not initial.
      node3 ?= node2->find_from_name_ns( name = 'blip' uri = namespace-a ).
      check node3 is not initial.
      me->fill_struct_from_attributes( exporting ip_element = node3 changing cp_structure = blip ).
      lv_relation_id = blip-embed.

      drawing_type = zcl_excel_drawing=>type_image.
    endif.

    node ?= io_anchor_element->find_from_name_ns( name = 'graphicFrame' uri = namespace-xdr ).
    if node is not initial.
      node2 ?= node->find_from_name_ns( name = 'nvGraphicFramePr' uri = namespace-xdr ).
      check node2 is not initial.
      node3 ?= node2->find_from_name_ns( name = 'cNvPr' uri = namespace-xdr ).
      check node3 is not initial.
      me->fill_struct_from_attributes( exporting ip_element = node3 changing cp_structure = cnvpr ).
      lv_title = cnvpr-name.

      node2 ?= node->find_from_name_ns( name = 'graphic' uri = namespace-a ).
      check node2 is not initial.
      node3 ?= node2->find_from_name_ns( name = 'graphicData' uri = namespace-a ).
      check node3 is not initial.
      node4 ?= node2->find_from_name_ns( name = 'chart' uri = namespace-c ).
      check node4 is not initial.
      me->fill_struct_from_attributes( exporting ip_element = node4 changing cp_structure = chart ).
      lv_relation_id = chart-id.

      drawing_type = zcl_excel_drawing=>type_chart.
    endif.

    lo_drawing = io_worksheet->excel->add_new_drawing(
                      ip_type  = drawing_type
                      ip_title = lv_title ).
    io_worksheet->add_drawing( lo_drawing ).

    lo_drawing->set_position2(
      exporting
        ip_from   = ls_upper
        ip_to     = ls_lower ).

    read table it_related_drawings into rel_drawing
          with key id = lv_relation_id.

    lo_drawing->set_media(
      exporting
        ip_media = rel_drawing-content
        ip_media_type = rel_drawing-file_ext
        ip_width = ls_size-width
        ip_height = ls_size-height ).

    if drawing_type = zcl_excel_drawing=>type_chart.
*  Begin fix for Issue #551
      data: lo_tmp_node_2                type ref to if_ixml_element.
      lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'pieChart' uri = namespace-c ).
      if lo_tmp_node_2 is not initial.
        lo_drawing->graph_type = zcl_excel_drawing=>c_graph_pie.
      else.
        lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'barChart' uri = namespace-c ).
        if lo_tmp_node_2 is not initial.
          lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.
        else.
          lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
          if lo_tmp_node_2 is not initial.
            lo_drawing->graph_type = zcl_excel_drawing=>c_graph_line.
          endif.
        endif.
      endif.
* End fix for issue #551
      "-------------Added by Alessandro Iannacci - Should load chart attributes
      lo_drawing->load_chart_attributes( rel_drawing-content_xml ).
    endif.

  endmethod.


  method load_dxf_styles.

    data: lo_styles_xml   type ref to if_ixml_document,
          lo_node_dxfs    type ref to if_ixml_element,

          lo_nodes_dxf    type ref to if_ixml_node_collection,
          lo_iterator_dxf type ref to if_ixml_node_iterator,
          lo_node_dxf     type ref to if_ixml_element,

          lv_dxf_count    type i.

    field-symbols: <ls_dxf_style> like line of mt_dxf_styles.

*--------------------------------------------------------------------*
* Look for dxfs-node
*--------------------------------------------------------------------*
    lo_styles_xml = me->get_ixml_from_zip_archive( iv_path ).
    lo_node_dxfs  = lo_styles_xml->find_from_name_ns( name = 'dxfs' uri = namespace-main ).
    check lo_node_dxfs is bound.


*--------------------------------------------------------------------*
* loop through all dxf-nodes and create style for each
*--------------------------------------------------------------------*
    lo_nodes_dxf ?= lo_node_dxfs->get_elements_by_tag_name_ns( name = 'dxf' uri = namespace-main ).
    lo_iterator_dxf = lo_nodes_dxf->create_iterator( ).
    lo_node_dxf ?= lo_iterator_dxf->get_next( ).
    while lo_node_dxf is bound.

      append initial line to mt_dxf_styles assigning <ls_dxf_style>.
      <ls_dxf_style>-dxf = lv_dxf_count. " We start counting at 0
      lv_dxf_count += 1.             " prepare next entry

      <ls_dxf_style>-guid = get_dxf_style_guid( io_ixml_dxf = lo_node_dxf
                                                io_excel    = io_excel ).
      lo_node_dxf ?= lo_iterator_dxf->get_next( ).

    endwhile.


  endmethod.


  method load_shared_strings.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support partial formatting of strings in cells
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: renaming variables to naming conventions
*          renaming variables to indicate what they are used for
*          aligning code
*          adding comments to explain what we are trying to achieve
*          rewriting code for better readibility
*--------------------------------------------------------------------*



    data:
      lo_shared_strings_xml type ref to if_ixml_document,
      lo_node_si            type ref to if_ixml_element,
      lo_node_si_child      type ref to if_ixml_element,
      lo_node_r_child_t     type ref to if_ixml_element,
      lo_node_r_child_rpr   type ref to if_ixml_element,
      lo_font               type ref to zcl_excel_style_font,
      ls_rtf                type zif_excel_data_decl=>zexcel_s_rtf,
      lv_current_offset     type int2,
      lv_tag_name           type string,
      lv_node_value         type string.

    field-symbols: <ls_shared_string>           like line of me->shared_strings.

*--------------------------------------------------------------------*

* §1  Parse shared strings file and get into internal table
*   So far I have encountered 2 ways how a string can be represented in the shared strings file
*   §1.1 - "simple" strings
*   §1.2 - rich text formatted strings

*     Following is an example how this file could be set up; 2 strings in simple formatting, 3rd string rich textformatted


*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <sst uniqueCount="6" count="6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <si>
*                <t>This is a teststring 1</t>
*            </si>
*            <si>
*                <t>This is a teststring 2</t>
*            </si>
*            <si>
*                <r>
*                  <t>T</t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color rgb="FFFF0000"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t xml:space="preserve">his is a </t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color theme="1"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t>teststring 3</t>
*                </r>
*            </si>
*        </sst>
*--------------------------------------------------------------------*

    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
                                                           is_normalizing = space ).  " NO!!! normalizing - otherwise leading blanks will be omitted and that is not really desired for the stringtable
    lo_node_si ?= lo_shared_strings_xml->find_from_name_ns( name = 'si' uri = namespace-main ).
    while lo_node_si is bound.

      append initial line to me->shared_strings assigning <ls_shared_string>.            " Each <si>-entry in the xml-file must lead to an entry in our stringtable
      lo_node_si_child ?= lo_node_si->get_first_child( ).
      if lo_node_si_child is bound.
        lv_tag_name = lo_node_si_child->get_name( ).
        if lv_tag_name = 't'.
*--------------------------------------------------------------------*
*   §1.1 - "simple" strings
*                Example:  see above
*--------------------------------------------------------------------*
          <ls_shared_string>-value = unescape_string_value( lo_node_si_child->get_value( ) ).
        else.
*--------------------------------------------------------------------*
*   §1.2 - rich text formatted strings
*       it is sufficient to strip the <t>...</t> tag from each <r>-tag and concatenate these
*       as long as rich text formatting is not supported (2do§1) ignore all info about formatting
*                Example:  see above
*--------------------------------------------------------------------*
          clear: lv_current_offset.
          while lo_node_si_child is bound.                                             " actually these children of <si> are <r>-tags
            lv_tag_name = lo_node_si_child->get_name( ).
            if lv_tag_name = 'r'.

              clear: ls_rtf.

              " extracting rich text formating data
              lo_node_r_child_rpr ?= lo_node_si_child->find_from_name_ns( name = 'rPr' uri = namespace-main ).
              if lo_node_r_child_rpr is bound.
                lo_font = load_style_font( lo_node_r_child_rpr ).
                ls_rtf-font = lo_font->get_structure( ).
              endif.
              ls_rtf-offset = lv_current_offset.
              " extract the <t>...</t> part of each <r>-tag
              lo_node_r_child_t ?= lo_node_si_child->find_from_name_ns( name = 't' uri = namespace-main ).
              if lo_node_r_child_t is bound.
                lv_node_value = unescape_string_value( lo_node_r_child_t->get_value( ) ).
                concatenate <ls_shared_string>-value lv_node_value into <ls_shared_string>-value respecting blanks.
                ls_rtf-length = strlen( lv_node_value ).

                if ls_rtf-length > 0.
                  lv_current_offset = strlen( <ls_shared_string>-value ).
                  append ls_rtf to <ls_shared_string>-rtf.
                endif.
              endif.
            endif.

            lo_node_si_child ?= lo_node_si_child->get_next( ).

          endwhile.
        endif.
      endif.

      lo_node_si ?= lo_node_si->get_next( ).
    endwhile.

  endmethod.


  method load_styles.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    types: begin of lty_xf,
             applyalignment    type string,
             applyborder       type string,
             applyfill         type string,
             applyfont         type string,
             applynumberformat type string,
             applyprotection   type string,
             borderid          type string,
             fillid            type string,
             fontid            type string,
             numfmtid          type string,
             pivotbutton       type string,
             quoteprefix       type string,
             xfid              type string,
           end of lty_xf.

    types: begin of lty_alignment,
             horizontal      type string,
             indent          type string,
             justifylastline type string,
             readingorder    type string,
             relativeindent  type string,
             shrinktofit     type string,
             textrotation    type string,
             vertical        type string,
             wraptext        type string,
           end of lty_alignment.

    types: begin of lty_protection,
             hidden type string,
             locked type string,
           end of lty_protection.

    data: lo_styles_xml                 type ref to if_ixml_document,
          lo_style                      type ref to zcl_excel_style,

          lt_num_formats                type zcl_excel_style_number_format=>t_num_formats,
          lt_fills                      type t_fills,
          lt_borders                    type t_borders,
          lt_fonts                      type t_fonts,

          ls_num_format                 type zcl_excel_style_number_format=>t_num_format,
          ls_fill                       type ref to zcl_excel_style_fill,
          ls_cell_border                type ref to zcl_excel_style_borders,
          ls_font                       type ref to zcl_excel_style_font,

          lo_node_cellxfs               type ref to if_ixml_element,
          lo_node_cellxfs_xf            type ref to if_ixml_element,
          lo_node_cellxfs_xf_alignment  type ref to if_ixml_element,
          lo_node_cellxfs_xf_protection type ref to if_ixml_element,

          lo_nodes_xf                   type ref to if_ixml_node_collection,
          lo_iterator_cellxfs           type ref to if_ixml_node_iterator,

          ls_xf                         type lty_xf,
          ls_alignment                  type lty_alignment,
          ls_protection                 type lty_protection,
          lv_index                      type i.

*--------------------------------------------------------------------*
* To build a complete style that fully describes how a cell looks like
* we need the various parts
* §1 - Numberformat
* §2 - Fillstyle
* §3 - Borders
* §4 - Font
* §5 - Alignment
* §6 - Protection

*          Following is an example how this part of a file could be set up
*              ...
*              parts with various formatinformation - see §1,§2,§3,§4
*              ...
*          <cellXfs count="26">
*              <xf numFmtId="0" borderId="0" fillId="0" fontId="0" xfId="0"/>
*              <xf numFmtId="0" borderId="0" fillId="2" fontId="0" xfId="0" applyFill="1"/>
*              <xf numFmtId="0" borderId="1" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="2" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="3" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="4" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="0" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              ...
*          </cellXfs>
*--------------------------------------------------------------------*

    lo_styles_xml = me->get_ixml_from_zip_archive( ip_path ).

*--------------------------------------------------------------------*
* The styles are build up from
* §1 number formats
* §2 fill styles
* §3 border styles
* §4 fonts
* These need to be read before we can try to build up a complete
* style that describes the look of a cell
*--------------------------------------------------------------------*
    lt_num_formats   = load_style_num_formats( lo_styles_xml ).   " §1
    lt_fills         = load_style_fills( lo_styles_xml ).         " §2
    lt_borders       = load_style_borders( lo_styles_xml ).       " §3
    lt_fonts         = load_style_fonts( lo_styles_xml ).         " §4

*--------------------------------------------------------------------*
* Now everything is prepared to build a "full" style
*--------------------------------------------------------------------*
    lo_node_cellxfs  = lo_styles_xml->find_from_name_ns( name = 'cellXfs' uri = namespace-main ).
    if lo_node_cellxfs is bound.
      lo_nodes_xf         = lo_node_cellxfs->get_elements_by_tag_name_ns( name = 'xf' uri = namespace-main ).
      lo_iterator_cellxfs = lo_nodes_xf->create_iterator( ).
      lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).
      while lo_node_cellxfs_xf is bound.

        lo_style = ip_excel->add_new_style( ).
        fill_struct_from_attributes( exporting
                                       ip_element   =  lo_node_cellxfs_xf
                                     changing
                                       cp_structure = ls_xf ).
*--------------------------------------------------------------------*
* §2 fill style
*--------------------------------------------------------------------*
        if ls_xf-applyfill = '1' and ls_xf-fillid is not initial.
          lv_index = ls_xf-fillid + 1.
          read table lt_fills into ls_fill index lv_index.
          if sy-subrc = 0.
            lo_style->fill = ls_fill.
          endif.
        endif.

*--------------------------------------------------------------------*
* §1 number format
*--------------------------------------------------------------------*
        if ls_xf-numfmtid is not initial.
          read table lt_num_formats into ls_num_format with table key id = ls_xf-numfmtid.
          if sy-subrc = 0.
            lo_style->number_format = ls_num_format-format.
          endif.
        endif.

*--------------------------------------------------------------------*
* §3 border style
*--------------------------------------------------------------------*
        if ls_xf-applyborder = '1' and ls_xf-borderid is not initial.
          lv_index = ls_xf-borderid + 1.
          read table lt_borders into ls_cell_border index lv_index.
          if sy-subrc = 0.
            lo_style->borders = ls_cell_border.
          endif.
        endif.

*--------------------------------------------------------------------*
* §4 font
*--------------------------------------------------------------------*
        if ls_xf-applyfont = '1' and ls_xf-fontid is not initial.
          lv_index = ls_xf-fontid + 1.
          read table lt_fonts into ls_font index lv_index.
          if sy-subrc = 0.
            lo_style->font = ls_font.
          endif.
        endif.

*--------------------------------------------------------------------*
* §5 - Alignment
*--------------------------------------------------------------------*
        lo_node_cellxfs_xf_alignment ?= lo_node_cellxfs_xf->find_from_name_ns( name = 'alignment' uri = namespace-main ).
        if lo_node_cellxfs_xf_alignment is bound.
          fill_struct_from_attributes( exporting
                                         ip_element   =  lo_node_cellxfs_xf_alignment
                                       changing
                                         cp_structure = ls_alignment ).
          if ls_alignment-horizontal is not initial.
            lo_style->alignment->horizontal = ls_alignment-horizontal.
          endif.

          if ls_alignment-vertical is not initial.
            lo_style->alignment->vertical = ls_alignment-vertical.
          endif.

          if ls_alignment-textrotation is not initial.
            lo_style->alignment->textrotation = ls_alignment-textrotation.
          endif.

          if ls_alignment-wraptext = '1' or ls_alignment-wraptext = 'true'.
            lo_style->alignment->wraptext = abap_true.
          endif.

          if ls_alignment-shrinktofit = '1' or ls_alignment-shrinktofit = 'true'.
            lo_style->alignment->shrinktofit = abap_true.
          endif.

          if ls_alignment-indent is not initial.
            lo_style->alignment->indent = ls_alignment-indent.
          endif.
        endif.

*--------------------------------------------------------------------*
* §6 - Protection
*--------------------------------------------------------------------*
        lo_node_cellxfs_xf_protection ?= lo_node_cellxfs_xf->find_from_name_ns( name = 'protection' uri = namespace-main ).
        if lo_node_cellxfs_xf_protection is bound.
          fill_struct_from_attributes( exporting
                                         ip_element   = lo_node_cellxfs_xf_protection
                                       changing
                                         cp_structure = ls_protection ).
          if ls_protection-locked = '1' or ls_protection-locked = 'true'.
            lo_style->protection->locked = zcl_excel_style_protection=>c_protection_locked.
          else.
            lo_style->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
          endif.

          if ls_protection-hidden = '1' or ls_protection-hidden = 'true'.
            lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_hidden.
          else.
            lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_unhidden.
          endif.

        endif.

        insert lo_style into table me->styles.

        lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).

      endwhile.
    endif.

  endmethod.


  method load_style_borders.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    data: lo_node_border      type ref to if_ixml_element,
          lo_node_bordertype  type ref to if_ixml_element,
          lo_node_bordercolor type ref to if_ixml_element,
          lo_cell_border      type ref to zcl_excel_style_borders,
          lo_border           type ref to zcl_excel_style_border,
          ls_color            type t_color.

*--------------------------------------------------------------------*
* We need a table of used borderformats to build up our styles
* §1    A cell has 4 outer borders and 2 diagonal "borders"
*       These borders can be formatted separately but the diagonal borders
*       are always being formatted the same
*       We'll parse through the <border>-tag for each of the bordertypes
* §2    and read the corresponding formatting information

*          Following is an example how this part of a file could be set up
*          <border diagonalDown="1">
*              <left style="mediumDashDotDot">
*                  <color rgb="FFFF0000"/>
*              </left>
*              <right/>
*              <top style="thick">
*                  <color rgb="FFFF0000"/>
*              </top>
*              <bottom style="thick">
*                  <color rgb="FFFF0000"/>
*              </bottom>
*              <diagonal style="thick">
*                  <color rgb="FFFF0000"/>
*              </diagonal>
*          </border>
*--------------------------------------------------------------------*
    lo_node_border ?= ip_xml->find_from_name_ns( name = 'border' uri = namespace-main ).
    while lo_node_border is bound.

      create object lo_cell_border.

*--------------------------------------------------------------------*
* Diagonal borderlines are formatted the equally.  Determine what kind of diagonal borders are present if any
*--------------------------------------------------------------------*
* DiagonalNone = 0
* DiagonalUp   = 1
* DiagonalDown = 2
* DiagonalBoth = 3
*--------------------------------------------------------------------*
      if lo_node_border->get_attribute( 'diagonalDown' ) is not initial.
        lo_cell_border->diagonal_mode += zcl_excel_style_borders=>c_diagonal_down.
      endif.

      if lo_node_border->get_attribute( 'diagonalUp' ) is not initial.
        lo_cell_border->diagonal_mode += zcl_excel_style_borders=>c_diagonal_up.
      endif.

      lo_node_bordertype ?= lo_node_border->get_first_child( ).
      while lo_node_bordertype is bound.
*--------------------------------------------------------------------*
* §1 Determine what kind of border we are talking about
*--------------------------------------------------------------------*
* Up, down, left, right, diagonal
*--------------------------------------------------------------------*
        create object lo_border.

        case lo_node_bordertype->get_name( ).

          when 'left'.
            lo_cell_border->left = lo_border.

          when 'right'.
            lo_cell_border->right = lo_border.

          when 'top'.
            lo_cell_border->top = lo_border.

          when 'bottom'.
            lo_cell_border->down = lo_border.

          when 'diagonal'.
            lo_cell_border->diagonal = lo_border.

        endcase.

*--------------------------------------------------------------------*
* §2 Read the border-formatting
*--------------------------------------------------------------------*
        lo_border->border_style = lo_node_bordertype->get_attribute( 'style' ).
        lo_node_bordercolor ?= lo_node_bordertype->find_from_name_ns( name = 'color' uri = namespace-main ).
        if lo_node_bordercolor is bound.
          fill_struct_from_attributes( exporting
                                         ip_element   =  lo_node_bordercolor
                                       changing
                                         cp_structure = ls_color ).

          lo_border->border_color-rgb = ls_color-rgb.
          if ls_color-indexed is not initial.
            lo_border->border_color-indexed = ls_color-indexed.
          endif.

          if ls_color-theme is not initial.
            lo_border->border_color-theme = ls_color-theme.
          endif.
          lo_border->border_color-tint = ls_color-tint.
        endif.

        lo_node_bordertype ?= lo_node_bordertype->get_next( ).

      endwhile.

      insert lo_cell_border into table ep_borders.

      lo_node_border ?= lo_node_border->get_next( ).

    endwhile.


  endmethod.


  method load_style_fills.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support gradientFill
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          commenting on problems/future enhancements/todos we already know of or should decide upon
*          adding comments to explain what we are trying to achieve
*          renaming variables to indicate what they are used for
*--------------------------------------------------------------------*
    data: lv_value           type string,
          lo_node_fill       type ref to if_ixml_element,
          lo_node_fill_child type ref to if_ixml_element,
          lo_node_bgcolor    type ref to if_ixml_element,
          lo_node_fgcolor    type ref to if_ixml_element,
          lo_node_stop       type ref to if_ixml_element,
          lo_fill            type ref to zcl_excel_style_fill,
          ls_color           type t_color.

*--------------------------------------------------------------------*
* We need a table of used fillformats to build up our styles

*          Following is an example how this part of a file could be set up
*          <fill>
*              <patternFill patternType="gray125"/>
*          </fill>
*          <fill>
*              <patternFill patternType="solid">
*                  <fgColor rgb="FFFFFF00"/>
*                  <bgColor indexed="64"/>
*              </patternFill>
*          </fill>
*--------------------------------------------------------------------*

    lo_node_fill ?= ip_xml->find_from_name_ns( name = 'fill' uri = namespace-main ).
    while lo_node_fill is bound.

      create object lo_fill.
      lo_node_fill_child ?= lo_node_fill->get_first_child( ).
      lv_value            = lo_node_fill_child->get_name( ).
      case lv_value.

*--------------------------------------------------------------------*
* Patternfill
*--------------------------------------------------------------------*
        when 'patternFill'.
          lo_fill->filltype = lo_node_fill_child->get_attribute( 'patternType' ).
*--------------------------------------------------------------------*
* Patternfill - background color
*--------------------------------------------------------------------*
          lo_node_bgcolor = lo_node_fill_child->find_from_name_ns( name = 'bgColor' uri = namespace-main ).
          if lo_node_bgcolor is bound.
            fill_struct_from_attributes( exporting
                                           ip_element   = lo_node_bgcolor
                                         changing
                                           cp_structure = ls_color ).

            lo_fill->bgcolor-rgb = ls_color-rgb.
            if ls_color-indexed is not initial.
              lo_fill->bgcolor-indexed = ls_color-indexed.
            endif.

            if ls_color-theme is not initial.
              lo_fill->bgcolor-theme = ls_color-theme.
            endif.
            lo_fill->bgcolor-tint = ls_color-tint.
          endif.

*--------------------------------------------------------------------*
* Patternfill - foreground color
*--------------------------------------------------------------------*
          lo_node_fgcolor = lo_node_fill->find_from_name_ns( name = 'fgColor' uri = namespace-main ).
          if lo_node_fgcolor is bound.
            fill_struct_from_attributes( exporting
                                           ip_element   = lo_node_fgcolor
                                         changing
                                           cp_structure = ls_color ).

            lo_fill->fgcolor-rgb = ls_color-rgb.
            if ls_color-indexed is not initial.
              lo_fill->fgcolor-indexed = ls_color-indexed.
            endif.

            if ls_color-theme is not initial.
              lo_fill->fgcolor-theme = ls_color-theme.
            endif.
            lo_fill->fgcolor-tint = ls_color-tint.
          endif.


*--------------------------------------------------------------------*
* gradientFill
*--------------------------------------------------------------------*
        when 'gradientFill'.
          lo_fill->gradtype-type   = lo_node_fill_child->get_attribute( 'type' ).
          lo_fill->gradtype-top    = lo_node_fill_child->get_attribute( 'top' ).
          lo_fill->gradtype-left   = lo_node_fill_child->get_attribute( 'left' ).
          lo_fill->gradtype-right  = lo_node_fill_child->get_attribute( 'right' ).
          lo_fill->gradtype-bottom = lo_node_fill_child->get_attribute( 'bottom' ).
          lo_fill->gradtype-degree = lo_node_fill_child->get_attribute( 'degree' ).
          free lo_node_stop.
          lo_node_stop ?= lo_node_fill_child->find_from_name_ns( name = 'stop' uri = namespace-main ).
          while lo_node_stop is bound.
            if lo_fill->gradtype-position1 is initial.
              lo_fill->gradtype-position1 = lo_node_stop->get_attribute( 'position' ).
              lo_node_bgcolor = lo_node_stop->find_from_name_ns( name = 'color' uri = namespace-main ).
              if lo_node_bgcolor is bound.
                fill_struct_from_attributes( exporting
                                                ip_element   = lo_node_bgcolor
                                              changing
                                                cp_structure = ls_color ).

                lo_fill->bgcolor-rgb = ls_color-rgb.
                if ls_color-indexed is not initial.
                  lo_fill->bgcolor-indexed = ls_color-indexed.
                endif.

                if ls_color-theme is not initial.
                  lo_fill->bgcolor-theme = ls_color-theme.
                endif.
                lo_fill->bgcolor-tint = ls_color-tint.
              endif.
            elseif lo_fill->gradtype-position2 is initial.
              lo_fill->gradtype-position2 = lo_node_stop->get_attribute( 'position' ).
              lo_node_fgcolor = lo_node_stop->find_from_name_ns( name = 'color' uri = namespace-main ).
              if lo_node_fgcolor is bound.
                fill_struct_from_attributes( exporting
                                               ip_element   = lo_node_fgcolor
                                             changing
                                               cp_structure = ls_color ).

                lo_fill->fgcolor-rgb = ls_color-rgb.
                if ls_color-indexed is not initial.
                  lo_fill->fgcolor-indexed = ls_color-indexed.
                endif.

                if ls_color-theme is not initial.
                  lo_fill->fgcolor-theme = ls_color-theme.
                endif.
                lo_fill->fgcolor-tint = ls_color-tint.
              endif.
            elseif lo_fill->gradtype-position3 is initial.
              lo_fill->gradtype-position3 = lo_node_stop->get_attribute( 'position' ).
              "BGColor is filled already with position 1 no need to check again
            endif.

            lo_node_stop ?= lo_node_stop->get_next( ).
          endwhile.

        when others.

      endcase.


      insert lo_fill into table ep_fills.

      lo_node_fill ?= lo_node_fill->get_next( ).

    endwhile.


  endmethod.


  method load_style_font.

    data: lo_node_font type ref to if_ixml_element,
          lo_node2     type ref to if_ixml_element,
          lo_font      type ref to zcl_excel_style_font,
          ls_color     type t_color.

    lo_node_font = io_xml_element.

    create object lo_font.
*--------------------------------------------------------------------*
*   Bold
*--------------------------------------------------------------------*
    if lo_node_font->find_from_name_ns( name = 'b' uri = namespace-main ) is bound.
      lo_font->bold = abap_true.
    endif.

*--------------------------------------------------------------------*
*   Italic
*--------------------------------------------------------------------*
    if lo_node_font->find_from_name_ns( name = 'i' uri = namespace-main ) is bound.
      lo_font->italic = abap_true.
    endif.

*--------------------------------------------------------------------*
*   Underline
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'u' uri = namespace-main ).
    if lo_node2 is bound.
      lo_font->underline      = abap_true.
      lo_font->underline_mode = lo_node2->get_attribute( 'val' ).
    endif.

*--------------------------------------------------------------------*
*   StrikeThrough
*--------------------------------------------------------------------*
    if lo_node_font->find_from_name_ns( name = 'strike' uri = namespace-main ) is bound.
      lo_font->strikethrough = abap_true.
    endif.

*--------------------------------------------------------------------*
*   Fontsize
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'sz' uri = namespace-main ).
    if lo_node2 is bound.
      lo_font->size = lo_node2->get_attribute( 'val' ).
    endif.

*--------------------------------------------------------------------*
*   Fontname
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'name' uri = namespace-main ).
    if lo_node2 is bound.
      lo_font->name = lo_node2->get_attribute( 'val' ).
    else.
      lo_node2 = lo_node_font->find_from_name_ns( name = 'rFont' uri = namespace-main ).
      if lo_node2 is bound.
        lo_font->name = lo_node2->get_attribute( 'val' ).
      endif.
    endif.

*--------------------------------------------------------------------*
*   Fontfamily
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'family' uri = namespace-main ).
    if lo_node2 is bound.
      lo_font->family = lo_node2->get_attribute( 'val' ).
    endif.

*--------------------------------------------------------------------*
*   Fontscheme
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'scheme' uri = namespace-main ).
    if lo_node2 is bound.
      lo_font->scheme = lo_node2->get_attribute( 'val' ).
    else.
      clear lo_font->scheme.
    endif.

*--------------------------------------------------------------------*
*   Fontcolor
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'color' uri = namespace-main ).
    if lo_node2 is bound.
      fill_struct_from_attributes( exporting
                                     ip_element   =  lo_node2
                                   changing
                                     cp_structure = ls_color ).
      lo_font->color-rgb = ls_color-rgb.
      if ls_color-indexed is not initial.
        lo_font->color-indexed = ls_color-indexed.
      endif.

      if ls_color-theme is not initial.
        lo_font->color-theme = ls_color-theme.
      endif.
      lo_font->color-tint = ls_color-tint.
    endif.

    ro_font = lo_font.

  endmethod.


  method load_style_fonts.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          removing unused variables
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    data: lo_node_font type ref to if_ixml_element,
          lo_font      type ref to zcl_excel_style_font.

*--------------------------------------------------------------------*
* We need a table of used fonts to build up our styles

*          Following is an example how this part of a file could be set up
*          <font>
*              <sz val="11"/>
*              <color theme="1"/>
*              <name val="Calibri"/>
*              <family val="2"/>
*              <scheme val="minor"/>
*          </font>
*--------------------------------------------------------------------*
    lo_node_font ?= ip_xml->find_from_name_ns( name = 'font' uri = namespace-main ).
    while lo_node_font is bound.

      lo_font = load_style_font( lo_node_font ).
      insert lo_font into table ep_fonts.

      lo_node_font ?= lo_node_font->get_next( ).

    endwhile.


  endmethod.


  method load_style_num_formats.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Explain gaps in predefined formats
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          adding comments to explain what we are trying to achieve
*          aligning code
*--------------------------------------------------------------------*
    data: lo_node_numfmt type ref to if_ixml_element,
          ls_num_format  type zcl_excel_style_number_format=>t_num_format.

*--------------------------------------------------------------------*
* We need a table of used numberformats to build up our styles
* there are two kinds of numberformats
* §1 built-in numberformats
* §2 and those that have been explicitly added by the createor of the excel-file
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1 built-in numberformats
*--------------------------------------------------------------------*
    ep_num_formats = zcl_excel_style_number_format=>mt_built_in_num_formats.

*--------------------------------------------------------------------*
* §2   Get non-internal numberformats that are found in the file explicitly

*         Following is an example how this part of a file could be set up
*         <numFmts count="1">
*             <numFmt formatCode="#,###,###,###,##0.00" numFmtId="164"/>
*         </numFmts>
*--------------------------------------------------------------------*
    lo_node_numfmt ?= ip_xml->find_from_name_ns( name = 'numFmt' uri = namespace-main ).
    while lo_node_numfmt is bound.

      clear ls_num_format.

      create object ls_num_format-format.
      ls_num_format-format->format_code = lo_node_numfmt->get_attribute( 'formatCode' ).
      ls_num_format-id                  = lo_node_numfmt->get_attribute( 'numFmtId' ).
      insert ls_num_format into table ep_num_formats.

      lo_node_numfmt                          ?= lo_node_numfmt->get_next( ).

    endwhile.


  endmethod.


  method load_theme.
    data theme type ref to zcl_excel_theme.
    data: lo_theme_xml type ref to if_ixml_document.
    create object theme.
    lo_theme_xml = me->get_ixml_from_zip_archive( iv_path ).
    theme->read_theme( io_theme_xml = lo_theme_xml  ).
    ip_excel->set_theme( io_theme = theme ).
  endmethod.


  method load_workbook.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Move macro-reading from zcl_excel_reader_xlsm to this class
*                autodetect existance of macro/vba content
*                Allow inputparameter to explicitly tell reader to ignore vba-content
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-10
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          removing unused variables
*          adding me-> where possible
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*          renaming i/o parameters:  previous input-parameter ip_path  holds a (full) filename and not a path   --> rename to iv_workbook_full_filename
*                                                             ip_excel renamed while being at it                --> rename to io_excel
*--------------------------------------------------------------------*
* issue #232   - Read worksheetstate hidden/veryHidden
*              - Stefan Schmoecker,                          2012-11-11
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-02
* changes:    correction in named ranges to correctly attach
*             sheetlocal names/ranges to the correct sheet
*--------------------------------------------------------------------*
* issue#284 - Copied formulae ignored when reading excelfile
*           - Stefan Schmoecker,                             2013-08-02
* changes:    initialize area to hold referenced formulaedata
*             after all worksheets have been read resolve formuae
*--------------------------------------------------------------------*

    constants: lcv_shared_strings             type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
               lcv_worksheet                  type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
               lcv_styles                     type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
               lcv_vba_project                type string value 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', "#EC NEEDED     for future incorporation of XLSM-reader
               lcv_theme                      type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - begin data declarations
*--------------------------------------------------------------------*
               lcv_worksheet_state_hidden     type string value 'hidden',
               lcv_worksheet_state_veryhidden type string value 'veryHidden'.
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - end data declarations
*--------------------------------------------------------------------*

    data:
      lv_path                    type string,
      lv_filename                type string,
      lv_full_filename           type string,

      lo_rels_workbook           type ref to if_ixml_document,
      lt_worksheets              type standard table of t_relationship with non-unique default key,
      lo_workbook                type ref to if_ixml_document,
      lv_workbook_index          type i,
      lv_worksheet_path          type string,
      ls_sheet                   type t_sheet,

      lo_node                    type ref to if_ixml_element,
      ls_relationship            type t_relationship,
      lo_worksheet               type ref to zcl_excel_worksheet,
      lo_range                   type ref to zcl_excel_range,
      lv_worksheet_title         type zif_excel_data_decl=>zexcel_sheet_title,
      lv_tabix                   type i,            " #235 - repeat rows/cols.  Needed to link defined name to correct worksheet

      ls_range                   type t_range,
      lv_range_value             type zif_excel_data_decl=>zexcel_range_value,
      lv_position_temp           type i,
*--------------------------------------------------------------------*
* #229: Set active worksheet - begin data declarations
*--------------------------------------------------------------------*
      lv_active_sheet_string     type string,
      lv_zexcel_active_worksheet type zif_excel_data_decl=>zexcel_active_worksheet,
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns  - added autofilter support while changing this section
      lo_autofilter              type ref to zcl_excel_autofilter,
      ls_area                    type zif_excel_data_decl=>zexcel_s_autofilter_area,
      lv_col_start_alpha         type zif_excel_data_decl=>zexcel_cell_column_alpha,
      lv_col_end_alpha           type zif_excel_data_decl=>zexcel_cell_column_alpha,
      lv_row_start               type zif_excel_data_decl=>zexcel_cell_row,
      lv_row_end                 type zif_excel_data_decl=>zexcel_cell_row,
      lv_regex                   type string,
      lv_range_value_1           type zif_excel_data_decl=>zexcel_range_value,
      lv_range_value_2           type zif_excel_data_decl=>zexcel_range_value.
*--------------------------------------------------------------------*
* #229: Set active worksheet - end data declarations
*--------------------------------------------------------------------*
    field-symbols: <worksheet> type t_relationship.


*--------------------------------------------------------------------*

* §1  Get the position of files related to this workbook
*         Usually this will be <root>/xl/workbook.xml
*         Thus the workbookroot will be <root>/xl/
*         The position of all related files will be given in file
*         <workbookroot>/_rels/<workbookfilename>.rels and their positions
*         be be given relative to the workbookroot

*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Id="rId6"/>
*            <Relationship Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Id="rId5"/>
*            <Relationship Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId1"/>
*            <Relationship Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId2"/>
*            <Relationship Target="worksheets/sheet3.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId3"/>
*            <Relationship Target="worksheets/sheet4.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId4"/>
*            <Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId7"/>
*        </Relationships>
*
* §2  Load data that is relevant to the complete workbook
*     Currently supported is:
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*   §2.2    Styles          - This holds all styles that are used in all worksheets
*   §2.3    Worksheets      - For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*   §2.4    [Themes]                - not supported
*   §2.5    [VBA (Macro)]           - supported in class zcl_excel_reader_xlsm but should be moved here and autodetect
*   ...
*
* §3  Some information is held in the workbookfile as well
*   §3.1    Names and order of of worksheets
*   §3.2    Active worksheet
*   §3.3    Defined names
*   ...
*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <fileVersion rupBuild="4506" lowestEdited="4" lastEdited="4" appName="xl"/>
*            <workbookPr defaultThemeVersion="124226"/>
*            <bookViews>
*                <workbookView activeTab="1" windowHeight="8445" windowWidth="19035" yWindow="120" xWindow="120"/>
*            </bookViews>
*            <sheets>
*                <sheet r:id="rId1" sheetId="1" name="Sheet1"/>
*                <sheet r:id="rId2" sheetId="2" name="Sheet2"/>
*                <sheet r:id="rId3" sheetId="3" name="Sheet3" state="hidden"/>
*                <sheet r:id="rId4" sheetId="4" name="Sheet4"/>
*            </sheets>
*            <definedNames/>
*            <calcPr calcId="125725"/>
*        </workbook>
*--------------------------------------------------------------------*

    clear me->mt_ref_formulae.                                                                              " ins issue#284

*--------------------------------------------------------------------*
* §1  Get the position of files related to this workbook
*     Entry into this method is with the filename of the workbook
*--------------------------------------------------------------------*
    split reverse( iv_workbook_full_filename ) at '/' into lv_filename lv_path.
    lv_path = reverse( lv_path ).
    lv_filename = reverse( lv_filename ).

    concatenate lv_path '_rels/' lv_filename '.rels'
        into lv_full_filename.
    lo_rels_workbook = me->get_ixml_from_zip_archive( lv_full_filename ).

    lo_node ?= lo_rels_workbook->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ). "#EC NOTEXT
    while lo_node is bound.

      me->fill_struct_from_attributes( exporting ip_element = lo_node changing cp_structure = ls_relationship ).

      case ls_relationship-type.

*--------------------------------------------------------------------*
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*--------------------------------------------------------------------*
        when lcv_shared_strings.
          concatenate lv_path ls_relationship-target
              into lv_full_filename.
          me->load_shared_strings( lv_full_filename ).

*--------------------------------------------------------------------*
*   §2.3    Worksheets
*           For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*           Shared strings and styles have to be present before we can start with creating the worksheets
*           thus we only store this information for use when parsing the workbookfile for sheetinformations
*--------------------------------------------------------------------*
        when lcv_worksheet.
          append ls_relationship to lt_worksheets.

*--------------------------------------------------------------------*
*   §2.2    Styles           - This holds the styles that are used in all worksheets
*--------------------------------------------------------------------*
        when lcv_styles.
          concatenate lv_path ls_relationship-target
              into lv_full_filename.
          me->load_styles( ip_path  = lv_full_filename
                           ip_excel = io_excel ).
          me->load_dxf_styles( iv_path  = lv_full_filename
                               io_excel = io_excel ).
        when lcv_theme.
          concatenate lv_path ls_relationship-target
              into lv_full_filename.
          me->load_theme(
                  exporting
                    iv_path  = lv_full_filename
                    ip_excel = io_excel   " Excel creator
                    ).
        when others.

      endcase.

      lo_node ?= lo_node->get_next( ).

    endwhile.

*--------------------------------------------------------------------*
* §3  Some information held in the workbookfile
*--------------------------------------------------------------------*
    lo_workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).

*--------------------------------------------------------------------*
*   §3.1    Names and order of of worksheets
*--------------------------------------------------------------------*
    lo_node           ?= lo_workbook->find_from_name_ns( name = 'sheet' uri = namespace-main ).
    lv_workbook_index  = 1.
    while lo_node is bound.

      me->fill_struct_from_attributes( exporting
                                         ip_element   = lo_node
                                       changing
                                         cp_structure = ls_sheet ).
*--------------------------------------------------------------------*
*       Create new worksheet in workbook with correct name
*--------------------------------------------------------------------*
      lv_worksheet_title = ls_sheet-name.
      if lv_workbook_index = 1.                                               " First sheet has been added automatically by creating io_excel
        lo_worksheet = io_excel->get_active_worksheet( ).
        lo_worksheet->set_title( lv_worksheet_title ).
      else.
        lo_worksheet = io_excel->add_new_worksheet( lv_worksheet_title ).
      endif.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - begin of coding
*       Set status hidden if necessary
*--------------------------------------------------------------------*
      case ls_sheet-state.

        when lcv_worksheet_state_hidden.
          lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_hidden.

        when lcv_worksheet_state_veryhidden.
          lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_veryhidden.

      endcase.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - end of coding
*--------------------------------------------------------------------*
*--------------------------------------------------------------------*
*       Load worksheetdata
*--------------------------------------------------------------------*
      read table lt_worksheets assigning <worksheet> with key id = ls_sheet-id.
      if sy-subrc = 0.
        <worksheet>-sheetid = ls_sheet-sheetid.                                "ins #235 - repeat rows/cols - needed to identify correct sheet
        <worksheet>-localsheetid = |{ lv_workbook_index - 1 }|.
        concatenate lv_path <worksheet>-target
            into lv_worksheet_path.
        me->load_worksheet( ip_path      = lv_worksheet_path
                            io_worksheet = lo_worksheet ).
        <worksheet>-worksheet = lo_worksheet.
      endif.

      lo_node ?= lo_node->get_next( ).
      lv_workbook_index += 1.

    endwhile.
    sort lt_worksheets by sheetid.                                              " needed for localSheetid -referencing

*--------------------------------------------------------------------*
*   #284: Set active worksheet - Resolve referenced formulae to
*                                explicit formulae those cells
*--------------------------------------------------------------------*
    me->resolve_referenced_formulae( ).
    " ins issue#284
*--------------------------------------------------------------------*
*   #229: Set active worksheet - begin coding
*   §3.2    Active worksheet
*--------------------------------------------------------------------*
    lv_zexcel_active_worksheet = 1.                                 " First sheet = active sheet if nothing else specified.
    lo_node ?=  lo_workbook->find_from_name_ns( name = 'workbookView' uri = namespace-main ).
    if lo_node is bound.
      lv_active_sheet_string = lo_node->get_attribute( 'activeTab' ).
      try.
          lv_zexcel_active_worksheet = lv_active_sheet_string + 1.  " EXCEL numbers the sheets from 0 onwards --> index into worksheettable is increased by one
        catch cx_sy_conversion_error. "#EC NO_HANDLER    - error here --> just use the default 1st sheet
      endtry.
    endif.
    io_excel->set_active_sheet_index( lv_zexcel_active_worksheet ).
*--------------------------------------------------------------------*
* #229: Set active worksheet - end coding
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
*   §3.3    Defined names
*           So far I have encountered these
*             - named ranges      - sheetlocal
*             - named ranges      - workbookglobal
*             - autofilters       - sheetlocal  ( special range )
*             - repeat rows/cols  - sheetlocal ( special range )
*
*--------------------------------------------------------------------*
    lo_node ?=  lo_workbook->find_from_name_ns( name = 'definedName' uri = namespace-main ).
    while lo_node is bound.

      clear lo_range.                                                                                       "ins issue #235 - repeat rows/cols
      me->fill_struct_from_attributes(  exporting
                                        ip_element   =  lo_node
                                       changing
                                         cp_structure = ls_range ).
      lv_range_value = lo_node->get_value( ).

      if ls_range-localsheetid is not initial.                                                              " issue #163+
*      READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY id = ls_range-localsheetid.                "del issue #235 - repeat rows/cols " issue #163+
*        lo_range = <worksheet>-worksheet->add_new_range( ).                                              "del issue #235 - repeat rows/cols " issue #163+
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - begin
*--------------------------------------------------------------------*
        read table lt_worksheets assigning <worksheet> with key localsheetid = ls_range-localsheetid.
        if sy-subrc = 0.
          case ls_range-name.

*--------------------------------------------------------------------*
* insert autofilters
*--------------------------------------------------------------------*
            when zcl_excel_autofilters=>c_autofilter.
              " begin Dennis Schaaf
              try.
                  zcl_excel_common=>convert_range2column_a_row( exporting i_range        = lv_range_value
                                                                importing e_column_start = lv_col_start_alpha
                                                                          e_column_end   = lv_col_end_alpha
                                                                          e_row_start    = ls_area-row_start
                                                                          e_row_end      = ls_area-row_end ).
                  ls_area-col_start = zcl_excel_common=>convert_column2int( lv_col_start_alpha ).
                  ls_area-col_end   = zcl_excel_common=>convert_column2int( lv_col_end_alpha ).
                  lo_autofilter = io_excel->add_new_autofilter( io_sheet = <worksheet>-worksheet ) .
                  lo_autofilter->set_filter_area( is_area = ls_area ).
                catch zcx_excel.
                  " we expected a range but it was not usable, so just ignore it
              endtry.
              " end Dennis Schaaf

*--------------------------------------------------------------------*
* repeat print rows/columns
*--------------------------------------------------------------------*
            when zif_excel_sheet_printsettings=>gcv_print_title_name.
              lo_range = <worksheet>-worksheet->add_new_range( ).
              lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
*--------------------------------------------------------------------*
* This might be a temporary solution.  Maybe ranges get be reworked
* to support areas consisting of multiple rectangles
* But for now just split the range into row and columnpart
*--------------------------------------------------------------------*
              clear:lv_range_value_1,
                    lv_range_value_2.
              if lv_range_value is initial.
* Empty --> nothing to do
              else.
                if lv_range_value(1) = `'`.  " Escaped
                  lv_regex = `^('[^']*')+![^,]*,`.
                else.
                  lv_regex = `^[^!]*![^,]*,`.
                endif.
* Split into two ranges if necessary
                find PCRE lv_regex in lv_range_value match length lv_position_temp.
                if sy-subrc = 0 and lv_position_temp > 0.
                  lv_range_value_2 = lv_range_value+lv_position_temp.
                  lv_position_temp -= 1.
                  lv_range_value_1 = lv_range_value(lv_position_temp).
                else.
                  lv_range_value_1 = lv_range_value.
                endif.
              endif.
* 1st range
              zcl_excel_common=>convert_range2column_a_row( exporting i_range            = lv_range_value_1
                                                                      i_allow_1dim_range = abap_true
                                                            importing e_column_start     = lv_col_start_alpha
                                                                      e_column_end       = lv_col_end_alpha
                                                                      e_row_start        = lv_row_start
                                                                      e_row_end          = lv_row_end ).
              if lv_col_start_alpha is not initial.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                      iv_columns_to   = lv_col_end_alpha ).
              endif.
              if lv_row_start is not initial.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                   iv_rows_to   = lv_row_end ).
              endif.

* 2nd range
              zcl_excel_common=>convert_range2column_a_row( exporting i_range            = lv_range_value_2
                                                                      i_allow_1dim_range = abap_true
                                                            importing e_column_start     = lv_col_start_alpha
                                                                      e_column_end       = lv_col_end_alpha
                                                                      e_row_start        = lv_row_start
                                                                      e_row_end          = lv_row_end ).
              if lv_col_start_alpha is not initial.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                      iv_columns_to   = lv_col_end_alpha ).
              endif.
              if lv_row_start is not initial.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                   iv_rows_to   = lv_row_end ).
              endif.

            when others.
              lo_range = <worksheet>-worksheet->add_new_range( ).

          endcase.
        endif.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - end
*--------------------------------------------------------------------*
      else.                                                                                                 " issue #163+
        lo_range = io_excel->add_new_range( ).                                                              " issue #163+
      endif.                                                                                                " issue #163+
*    lo_range = ip_excel->add_new_range( ).                                                               " issue #163-
      if lo_range is bound.                                                                                 "ins issue #235 - repeat rows/cols
        lo_range->name = ls_range-name.
        lo_range->set_range_value( lv_range_value ).
      endif.                                                                                                "ins issue #235 - repeat rows/cols
      lo_node ?= lo_node->get_next( ).

    endwhile.

  endmethod.


  method load_worksheet.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Header/footer
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,
*              - ...
* changes: renaming variables to naming conventions
*          aligning code                                            (started)
*          add a list of open ToDos here
*          adding comments to explain what we are trying to achieve (started)
*--------------------------------------------------------------------*
* issue #345 - Dump on small pagemargins
*              Took the chance to modularize this very long method
*              by extracting the code that needed correction into
*              own method ( load_worksheet_pagemargins )
*--------------------------------------------------------------------*
    types: begin of lty_cell,
             r type string,
             t type string,
             s type string,
           end of lty_cell.

    types: begin of lty_column,
             min          type string,
             max          type string,
             width        type f,
             customwidth  type string,
             style        type string,
             bestfit      type string,
             collapsed    type string,
             hidden       type string,
             outlinelevel type string,
           end of lty_column.

    types: begin of lty_sheetview,
             showgridlines            type zif_excel_data_decl=>zexcel_show_gridlines,
             tabselected              type string,
             zoomscale                type string,
             zoomscalenormal          type string,
             zoomscalepagelayoutview  type string,
             zoomscalesheetlayoutview type string,
             workbookviewid           type string,
             showrowcolheaders        type string,
             righttoleft              type string,
             topleftcell              type string,
           end of lty_sheetview.

    types: begin of lty_mergecell,
             ref type string,
           end of lty_mergecell.

    types: begin of lty_row,
             r            type string,
             customheight type string,
             ht           type f,
             spans        type string,
             thickbot     type string,
             customformat type string,
             thicktop     type string,
             collapsed    type string,
             hidden       type string,
             outlinelevel type string,
           end of lty_row.

    types: begin of lty_page_setup,
             id          type string,
             orientation type string,
             scale       type string,
             fittoheight type string,
             fittowidth  type string,
             papersize   type string,
             paperwidth  type string,
             paperheight type string,
           end of lty_page_setup.

    types: begin of lty_sheetformatpr,
             customheight     type string,
             defaultrowheight type string,
             customwidth      type string,
             defaultcolwidth  type string,
           end of lty_sheetformatpr.

    types: begin of lty_headerfooter,
             alignwithmargins type string,
             differentoddeven type string,
           end of lty_headerfooter.

    types: begin of lty_tabcolor,
             rgb   type string,
             theme type string,
           end of lty_tabcolor.

    types: begin of lty_datavalidation,
             type             type zif_excel_data_decl=>zexcel_data_val_type,
             allowblank       type abap_boolean,
             showinputmessage type abap_boolean,
             showerrormessage type abap_boolean,
             showdropdown     type abap_boolean,
             operator         type zif_excel_data_decl=>zexcel_data_val_operator,
             formula1         type zif_excel_data_decl=>zexcel_validation_formula1,
             formula2         type zif_excel_data_decl=>zexcel_validation_formula1,
             sqref            type string,
             cell_column      type zif_excel_data_decl=>zexcel_cell_column_alpha,
             cell_column_to   type zif_excel_data_decl=>zexcel_cell_column_alpha,
             cell_row         type zif_excel_data_decl=>zexcel_cell_row,
             cell_row_to      type zif_excel_data_decl=>zexcel_cell_row,
             error            type string,
             errortitle       type string,
             prompt           type string,
             prompttitle      type string,
             errorstyle       type zif_excel_data_decl=>zexcel_data_val_error_style,
           end of lty_datavalidation.



    constants: lc_xml_attr_true     type string value 'true',
               lc_xml_attr_true_int type string value '1',
               lc_rel_drawing       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
               lc_rel_hyperlink     type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
               lc_rel_comments      type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
               lc_rel_printer       type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'.
    constants lc_rel_table type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'.

    data: lo_ixml_worksheet           type ref to if_ixml_document,
          lo_ixml_cells               type ref to if_ixml_node_collection,
          lo_ixml_iterator            type ref to if_ixml_node_iterator,
          lo_ixml_iterator2           type ref to if_ixml_node_iterator,
          lo_ixml_row_elem            type ref to if_ixml_element,
          lo_ixml_cell_elem           type ref to if_ixml_element,
          ls_cell                     type lty_cell,
          lv_index                    type i,
          lv_index_temp               type i,
          lo_ixml_value_elem          type ref to if_ixml_element,
          lo_ixml_formula_elem        type ref to if_ixml_element,
          lv_cell_value               type zif_excel_data_decl=>zexcel_cell_value,
          lv_cell_formula             type zif_excel_data_decl=>zexcel_cell_formula,
          lv_cell_column              type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_cell_row                 type zif_excel_data_decl=>zexcel_cell_row,
          lo_excel_style              type ref to zcl_excel_style,
          lv_style_guid               type zif_excel_data_decl=>zexcel_cell_style,

          lo_ixml_imension_elem       type ref to if_ixml_element, "#+234
          lv_dimension_range          type string,  "#+234

          lo_ixml_sheetview_elem      type ref to if_ixml_element,
          ls_sheetview                type lty_sheetview,
          lo_ixml_pane_elem           type ref to if_ixml_element,
          ls_excel_pane               type zif_excel_data_decl=>zexcel_pane,
          lv_pane_cell_row            type zif_excel_data_decl=>zexcel_cell_row,
          lv_pane_cell_col_a          type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_pane_cell_col            type zif_excel_data_decl=>zexcel_cell_column,

          lo_ixml_mergecells          type ref to if_ixml_node_collection,
          lo_ixml_mergecell_elem      type ref to if_ixml_element,
          ls_mergecell                type lty_mergecell,
          lv_merge_column_start       type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_merge_column_end         type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_merge_row_start          type zif_excel_data_decl=>zexcel_cell_row,
          lv_merge_row_end            type zif_excel_data_decl=>zexcel_cell_row,

          lo_ixml_sheetformatpr_elem  type ref to if_ixml_element,
          ls_sheetformatpr            type lty_sheetformatpr,
          lv_height                   type f,

          lo_ixml_headerfooter_elem   type ref to if_ixml_element,
          ls_headerfooter             type lty_headerfooter,
          ls_odd_header               type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_odd_footer               type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_even_header              type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_even_footer              type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          lo_ixml_hf_value_elem       type ref to if_ixml_element,

          lo_ixml_pagesetup_elem      type ref to if_ixml_element,
          lo_ixml_sheetpr             type ref to if_ixml_element,
          lv_fit_to_page              type string,
          ls_pagesetup                type lty_page_setup,

          lo_ixml_columns             type ref to if_ixml_node_collection,
          lo_ixml_column_elem         type ref to if_ixml_element,
          ls_column                   type lty_column,
          lv_column_alpha             type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lo_column                   type ref to zcl_excel_column,
          lv_outline_level            type int4,

          lo_ixml_tabcolor            type ref to if_ixml_element,
          ls_tabcolor                 type lty_tabcolor,
          ls_excel_s_tabcolor         type zif_excel_data_decl=>zexcel_s_tabcolor,

          lo_ixml_rows                type ref to if_ixml_node_collection,
          ls_row                      type lty_row,
          lv_max_col                  type i,     "for use with SPANS element
*              lv_min_col                     TYPE i,     "for use with SPANS element                    " not in use currently
          lv_max_col_s                type c length 10,     "for use with SPANS element
          lv_min_col_s                type c length 10,     "for use with SPANS element
          lo_row                      type ref to zcl_excel_row,
*---    End of current code aligning -------------------------------------------------------------------

          lv_path                     type string,
          lo_ixml_node                type ref to if_ixml_element,
          ls_relationship             type t_relationship,
          lo_ixml_rels_worksheet      type ref to if_ixml_document,
          lv_rels_worksheet_path      type string,
          lv_stripped_name            type string,
          lv_dirname                  type string,

          lt_external_hyperlinks      type gtt_external_hyperlinks,
          ls_external_hyperlink       like line of lt_external_hyperlinks,

          lo_ixml_datavalidations     type ref to if_ixml_node_collection,
          lo_ixml_datavalidation_elem type ref to if_ixml_element,
          ls_datavalidation           type lty_datavalidation,
          lo_data_validation          type ref to zcl_excel_data_validation,
          lv_datavalidation_range     type string,
          lt_datavalidation_range     type table of string,
          lt_rtf                      type zif_excel_data_decl=>zexcel_t_rtf,
          ex                          type ref to cx_root.
    data lt_tables type t_tables.
    data ls_table type t_table.

    field-symbols:
      <ls_shared_string> type t_shared_string.

*--------------------------------------------------------------------*
* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets
*--------------------------------------------------------------------*

    " Read Workbook Relationships
    split reverse( ip_path ) at '/' into lv_stripped_name lv_dirname.
    lv_dirname = reverse( lv_dirname ).
    lv_stripped_name = reverse( lv_stripped_name ).
    concatenate lv_dirname '_rels/' lv_stripped_name '.rels'
      into lv_rels_worksheet_path.
    try.                                                                          " +#222  _rels/xxx.rels might not be present.  If not found there can be no drawings --> just ignore this section
        lo_ixml_rels_worksheet = me->get_ixml_from_zip_archive( lv_rels_worksheet_path ).
        lo_ixml_node ?= lo_ixml_rels_worksheet->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
      catch zcx_excel.                            "#EC NO_HANDLER +#222
        " +#222   No errorhandling necessary - node will be unbound if error occurs
    endtry.                                                   " +#222
    while lo_ixml_node is bound.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_node
                                   changing
                                     cp_structure = ls_relationship ).
      concatenate lv_dirname ls_relationship-target into lv_path.
      lv_path = resolve_path( lv_path ).

      case ls_relationship-type.
        when lc_rel_drawing.
          " Read Drawings
* Issue # 339       Not all drawings are in the path mentioned below.
*                   Some Excel elements like textfields (which we don't support ) have a drawing-part in the relationsships
*                   but no "xl/drawings/_rels/drawing____.xml.rels" part.
*                   Since we don't support these there is no need to read them.  Catching exceptions thrown
*                   in the "load_worksheet_drawing" shouldn't lead to an abortion of the reading
          try.
              me->load_worksheet_drawing( ip_path      = lv_path
                                        io_worksheet = io_worksheet ).
            catch zcx_excel. "--> then ignore it
          endtry.

        when lc_rel_printer.
          " Read Printer settings

        when lc_rel_hyperlink.
          move-corresponding ls_relationship to ls_external_hyperlink.
          insert ls_external_hyperlink into table lt_external_hyperlinks.

        when lc_rel_comments.
          try.
              me->load_comments( ip_path      = lv_path
                                 io_worksheet = io_worksheet ).
            catch zcx_excel.
          endtry.

        when lc_rel_table.
          move-corresponding ls_relationship to ls_table.
          insert ls_table into table lt_tables.

        when others.
      endcase.

      lo_ixml_node ?= lo_ixml_node->get_next( ).
    endwhile.


    lo_ixml_worksheet = me->get_ixml_from_zip_archive( ip_path ).


    lo_ixml_tabcolor ?= lo_ixml_worksheet->find_from_name_ns( name = 'tabColor' uri = namespace-main ).
    if lo_ixml_tabcolor is bound.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_tabcolor
                                  changing
                                    cp_structure = ls_tabcolor ).
* Theme not supported yet
      if ls_tabcolor-rgb is not initial.
        ls_excel_s_tabcolor-rgb = ls_tabcolor-rgb.
        io_worksheet->set_tabcolor( ls_excel_s_tabcolor ).
      endif.
    endif.

    " Read tables (must be done before loading sheet contents)
    try.
        me->load_worksheet_tables( io_ixml_worksheet = lo_ixml_worksheet
                                   io_worksheet      = io_worksheet
                                   iv_dirname        = lv_dirname
                                   it_tables         = lt_tables ).
      catch zcx_excel. " Ignore reading errors - pass everything we were able to identify
    endtry.

    " Sheet contents
    lo_ixml_rows = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'row' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_rows->create_iterator( ).
    lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_row_elem is bound.

      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_row_elem
                                   changing
                                     cp_structure = ls_row ).
      split ls_row-spans at ':' into lv_min_col_s lv_max_col_s.
      lv_index = lv_max_col_s.
      if lv_index > lv_max_col.
        lv_max_col = lv_index.
      endif.
      lv_cell_row = ls_row-r.
      lv_cell_column = ''.
      lo_row = io_worksheet->get_row( lv_cell_row ).
      if ls_row-customheight = '1'.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_true ).
      elseif ls_row-ht > 0.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_false ).
      endif.

      if   ls_row-collapsed = lc_xml_attr_true
        or ls_row-collapsed = lc_xml_attr_true_int.
        lo_row->set_collapsed( abap_true ).
      endif.

      if   ls_row-hidden = lc_xml_attr_true
        or ls_row-hidden = lc_xml_attr_true_int.
        lo_row->set_visible( abap_false ).
      endif.

      if ls_row-outlinelevel > ''.
*        outline_level = condense( row-outlineLevel ).  "For basis 7.02 and higher
        condense  ls_row-outlinelevel.
        lv_outline_level = ls_row-outlinelevel.
        if lv_outline_level > 0.
          lo_row->set_outline_level( lv_outline_level ).
        endif.
      endif.

      lo_ixml_cells = lo_ixml_row_elem->get_elements_by_tag_name_ns( name = 'c' uri = namespace-main ).
      lo_ixml_iterator2 = lo_ixml_cells->create_iterator( ).
      lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
      while lo_ixml_cell_elem is bound.
        clear: lv_cell_value,
               lv_cell_formula,
               lv_style_guid,
               lt_rtf.

        fill_struct_from_attributes( exporting ip_element = lo_ixml_cell_elem changing cp_structure = ls_cell ).

        " Determine the column number
        if ls_cell-r is not initial.
          " Note that the row should remain unchanged = the one defined by <row>
          " i.e. in <row r="1"...><c r="A1" s="2"><v>..., ls_cell-r would be "A1",
          "      the "1" of A1 should always be equal to the "1" of <row r="1"...
          zcl_excel_common=>convert_columnrow2column_a_row( exporting
                                                              i_columnrow = ls_cell-r
                                                            importing
                                                              e_column    = lv_cell_column
                                                              e_row       = lv_cell_row ).
        else.
          " The column is the column after the last cell previously initialized in the same row.
          " NB: the row is unchanged = the one defined by <row> e.g. "1" in <row r="1"...><c r="" s="2"><v>...
          if lv_cell_column is initial.
            lv_cell_column = 'A'.
          else.
            lv_cell_column = zcl_excel_common=>convert_column2alpha( zcl_excel_common=>convert_column2int( lv_cell_column ) + 1 ).
          endif.
        endif.

        lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'v' uri = namespace-main ).

        case ls_cell-t.
          when 's'. " String values are stored as index in shared string table
            if lo_ixml_value_elem is bound.
              lv_index = lo_ixml_value_elem->get_value( ) + 1.
              read table shared_strings assigning <ls_shared_string> index lv_index.
              if sy-subrc = 0.
                lv_cell_value = <ls_shared_string>-value.
                lt_rtf = <ls_shared_string>-rtf.
              endif.
            endif.
          when 'inlineStr'. " inlineStr values are kept in special node
            lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'is' uri = namespace-main ).
            if lo_ixml_value_elem is bound.
              lv_cell_value = lo_ixml_value_elem->get_value( ).
            endif.
          when others. "other types are stored directly
            if lo_ixml_value_elem is bound.
              lv_cell_value = lo_ixml_value_elem->get_value( ).
            endif.
        endcase.

        clear lv_style_guid.
        "read style based on index
        if ls_cell-s is not initial.
          lv_index = ls_cell-s + 1.
          read table styles into lo_excel_style index lv_index.
          if sy-subrc = 0.
            lv_style_guid = lo_excel_style->get_guid( ).
          endif.
        endif.

        lo_ixml_formula_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'f' uri = namespace-main ).
        if lo_ixml_formula_elem is bound.
          lv_cell_formula = lo_ixml_formula_elem->get_value( ).
*--------------------------------------------------------------------*
* Begin of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
          data: begin of ls_formula_attributes,
                  ref type string,
                  si  type i,
                  t   type string,
                end of ls_formula_attributes,
                ls_ref_formula type ty_ref_formulae.

          fill_struct_from_attributes( exporting ip_element = lo_ixml_formula_elem changing cp_structure = ls_formula_attributes ).
          if ls_formula_attributes-t = 'shared'.

            try.
                clear ls_ref_formula.
                ls_ref_formula-sheet     = io_worksheet.
                ls_ref_formula-row       = lv_cell_row.
                ls_ref_formula-column    = zcl_excel_common=>convert_column2int( lv_cell_column ).
                ls_ref_formula-si        = ls_formula_attributes-si.
                ls_ref_formula-ref       = ls_formula_attributes-ref.
                ls_ref_formula-formula   = lv_cell_formula.
                insert ls_ref_formula into table me->mt_ref_formulae.
              catch cx_root into ex.
                raise exception type zcx_excel
                  exporting
                    previous = ex.
            endtry.
          endif.
*--------------------------------------------------------------------*
* End of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
        endif.

        if   lv_cell_value    is not initial
          or lv_cell_formula  is not initial
          or lv_style_guid    is not initial.
          io_worksheet->set_cell( ip_column     = lv_cell_column  " cell_elem Column
                                  ip_row        = lv_cell_row     " cell_elem row_elem
                                  ip_value      = lv_cell_value   " cell_elem Value
                                  ip_formula    = lv_cell_formula
                                  ip_data_type  = ls_cell-t
                                  ip_style      = lv_style_guid
                                  it_rtf        = lt_rtf ).
        endif.
        lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
      endwhile.
      lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
    endwhile.

*--------------------------------------------------------------------*
*#234 - column width not read correctly - begin of coding
*       reason - libre office doesn't use SPAN in row - definitions
*--------------------------------------------------------------------*
    if lv_max_col = 0.
      lo_ixml_imension_elem = lo_ixml_worksheet->find_from_name_ns( name = 'dimension' uri = namespace-main ).
      if lo_ixml_imension_elem is bound.
        lv_dimension_range = lo_ixml_imension_elem->get_attribute( 'ref' ).
        if lv_dimension_range cs ':'.
          replace PCRE  '\D+\d+:(\D+)\d+' in lv_dimension_range with '$1'.  " Get max column
        else.
          replace PCRE  '(\D+)\d+' in lv_dimension_range with '$1'.  " Get max column
        endif.
        lv_max_col = zcl_excel_common=>convert_column2int( lv_dimension_range ).
      endif.
    endif.
*--------------------------------------------------------------------*
*#234 - column width not read correctly - end of coding
*--------------------------------------------------------------------*

    "Get the customized column width
    lo_ixml_columns = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'col' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_columns->create_iterator( ).
    lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_column_elem is bound.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_column_elem
                                   changing
                                     cp_structure = ls_column ).
      lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
      if   ls_column-customwidth   = lc_xml_attr_true
        or ls_column-customwidth   = lc_xml_attr_true_int
        or ls_column-bestfit       = lc_xml_attr_true
        or ls_column-bestfit       = lc_xml_attr_true_int
        or ls_column-collapsed     = lc_xml_attr_true
        or ls_column-collapsed     = lc_xml_attr_true_int
        or ls_column-hidden        = lc_xml_attr_true
        or ls_column-hidden        = lc_xml_attr_true_int
        or ls_column-outlinelevel  > ''
        or ls_column-style         > ''.
        lv_index = ls_column-min.
        while lv_index <= ls_column-max and lv_index <= lv_max_col.

          lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_index ).
          lo_column =  io_worksheet->get_column( lv_column_alpha ).

          if   ls_column-customwidth = lc_xml_attr_true
            or ls_column-customwidth = lc_xml_attr_true_int
            or ls_column-width       is not initial.          "+#234
            lo_column->set_width( ls_column-width ).
          endif.

          if   ls_column-bestfit = lc_xml_attr_true
            or ls_column-bestfit = lc_xml_attr_true_int.
            lo_column->set_auto_size( abap_true ).
          endif.

          if   ls_column-collapsed = lc_xml_attr_true
            or ls_column-collapsed = lc_xml_attr_true_int.
            lo_column->set_collapsed( abap_true ).
          endif.

          if   ls_column-hidden = lc_xml_attr_true
            or ls_column-hidden = lc_xml_attr_true_int.
            lo_column->set_visible( abap_false ).
          endif.

          if ls_column-outlinelevel > ''.
            condense ls_column-outlinelevel.
            lv_outline_level = ls_column-outlinelevel.
            if lv_outline_level > 0.
              lo_column->set_outline_level( lv_outline_level ).
            endif.
          endif.

          if ls_column-style > ''.
            lv_index_temp = ls_column-style + 1.
            read table styles into lo_excel_style index lv_index_temp.
            data: dummy_zexcel_cell_style type zif_excel_data_decl=>zexcel_cell_style.
            dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
            lo_column->set_column_style_by_guid( dummy_zexcel_cell_style ).
          endif.

          lv_index += 1.
        endwhile.
      endif.

* issue #367 - hide columns from
      if ls_column-max = zcl_excel_common=>c_excel_sheet_max_col.     " Max = very right column
        if ( ls_column-hidden = lc_xml_attr_true
          or ls_column-hidden = lc_xml_attr_true_int ) " all hidden
          and ls_column-min > 0.
          io_worksheet->zif_excel_sheet_properties~hide_columns_from = zcl_excel_common=>convert_column2alpha( ls_column-min ).
        elseif ls_column-style > ''.
          lv_index_temp = ls_column-style + 1.
          read table styles into lo_excel_style index lv_index_temp.
          dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
* Set style for remaining columns
          io_worksheet->zif_excel_sheet_properties~set_style( dummy_zexcel_cell_style ).
        endif.
      endif.


    endwhile.

    "Now we need to get information from the sheetView node
    lo_ixml_sheetview_elem = lo_ixml_worksheet->find_from_name_ns( name = 'sheetView' uri = namespace-main ).
    fill_struct_from_attributes( exporting ip_element = lo_ixml_sheetview_elem changing cp_structure = ls_sheetview ).
    if ls_sheetview-showgridlines is initial or
       ls_sheetview-showgridlines = lc_xml_attr_true or
       ls_sheetview-showgridlines = lc_xml_attr_true_int.
      "If the attribute is not specified or set to true, we will show grid lines
      ls_sheetview-showgridlines = abap_true.
    else.
      ls_sheetview-showgridlines = abap_false.
    endif.
    io_worksheet->set_show_gridlines( ls_sheetview-showgridlines ).
    if ls_sheetview-righttoleft = lc_xml_attr_true
        or ls_sheetview-righttoleft = lc_xml_attr_true_int.
      io_worksheet->zif_excel_sheet_properties~set_right_to_left( abap_true ).
    endif.
    io_worksheet->zif_excel_sheet_properties~zoomscale                 = ls_sheetview-zoomscale.
    io_worksheet->zif_excel_sheet_properties~zoomscale_normal          = ls_sheetview-zoomscalenormal.
    io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview  = ls_sheetview-zoomscalepagelayoutview.
    io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = ls_sheetview-zoomscalesheetlayoutview.
    if ls_sheetview-topleftcell is not initial.
      io_worksheet->set_sheetview_top_left_cell( ls_sheetview-topleftcell ).
    endif.

    "Add merge cell information
    lo_ixml_mergecells = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'mergeCell' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_mergecells->create_iterator( ).
    lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_mergecell_elem is bound.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_mergecell_elem
                                   changing
                                     cp_structure = ls_mergecell ).
      zcl_excel_common=>convert_range2column_a_row( exporting
                                                      i_range = ls_mergecell-ref
                                                    importing
                                                      e_column_start = lv_merge_column_start
                                                      e_column_end   = lv_merge_column_end
                                                      e_row_start    = lv_merge_row_start
                                                      e_row_end      = lv_merge_row_end ).
      lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
      io_worksheet->set_merge( exporting
                                 ip_column_start = lv_merge_column_start
                                 ip_column_end   = lv_merge_column_end
                                 ip_row          = lv_merge_row_start
                                 ip_row_to       = lv_merge_row_end ).
    endwhile.

    " read sheet format properties
    lo_ixml_sheetformatpr_elem = lo_ixml_worksheet->find_from_name_ns( name = 'sheetFormatPr' uri = namespace-main ).
    if lo_ixml_sheetformatpr_elem is not initial.
      fill_struct_from_attributes( exporting ip_element = lo_ixml_sheetformatpr_elem changing cp_structure = ls_sheetformatpr ).
      if ls_sheetformatpr-customheight = '1'.
        lv_height = ls_sheetformatpr-defaultrowheight.
        lo_row = io_worksheet->get_default_row( ).
        lo_row->set_row_height( lv_height ).
      endif.

      " TODO...  column
    endif.

    " Read in page margins
    me->load_worksheet_pagemargins( exporting
                                      io_ixml_worksheet = lo_ixml_worksheet
                                      io_worksheet      = io_worksheet ).

* FitToPage
    lo_ixml_sheetpr ?=  lo_ixml_worksheet->find_from_name_ns( name = 'pageSetUpPr' uri = namespace-main ).
    if lo_ixml_sheetpr is bound.

      lv_fit_to_page = lo_ixml_sheetpr->get_attribute_ns( 'fitToPage' ).
      if lv_fit_to_page is not initial.
        io_worksheet->sheet_setup->fit_to_page = 'X'.
      endif.
    endif.
    " Read in page setup
    lo_ixml_pagesetup_elem = lo_ixml_worksheet->find_from_name_ns( name = 'pageSetup' uri = namespace-main ).
    if lo_ixml_pagesetup_elem is not initial.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_pagesetup_elem
                                   changing
                                     cp_structure = ls_pagesetup ).
      io_worksheet->sheet_setup->orientation = ls_pagesetup-orientation.
      io_worksheet->sheet_setup->scale = ls_pagesetup-scale.
      io_worksheet->sheet_setup->paper_size = ls_pagesetup-papersize.
      io_worksheet->sheet_setup->paper_height = ls_pagesetup-paperheight.
      io_worksheet->sheet_setup->paper_width = ls_pagesetup-paperwidth.
      if io_worksheet->sheet_setup->fit_to_page = 'X'.
        if ls_pagesetup-fittowidth is not initial.
          io_worksheet->sheet_setup->fit_to_width = ls_pagesetup-fittowidth.
        else.
          io_worksheet->sheet_setup->fit_to_width = 1.  " Default if not given - Excel doesn't write this to xml
        endif.
        if ls_pagesetup-fittoheight is not initial.
          io_worksheet->sheet_setup->fit_to_height = ls_pagesetup-fittoheight.
        else.
          io_worksheet->sheet_setup->fit_to_height = 1. " Default if not given - Excel doesn't write this to xml
        endif.
      endif.
    endif.



    " Read header footer
    lo_ixml_headerfooter_elem = lo_ixml_worksheet->find_from_name_ns( name = 'headerFooter' uri = namespace-main ).
    if lo_ixml_headerfooter_elem is not initial.
      fill_struct_from_attributes( exporting ip_element = lo_ixml_headerfooter_elem changing cp_structure = ls_headerfooter ).
      io_worksheet->sheet_setup->diff_oddeven_headerfooter = ls_headerfooter-differentoddeven.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'oddFooter' uri = namespace-main ).
      if lo_ixml_hf_value_elem is not initial.
        ls_odd_footer-left_value = lo_ixml_hf_value_elem->get_value( ).
      endif.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'oddHeader' uri = namespace-main ).
      if lo_ixml_hf_value_elem is not initial.
        ls_odd_header-left_value = lo_ixml_hf_value_elem->get_value( ).
      endif.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'evenFooter' uri = namespace-main ).
      if lo_ixml_hf_value_elem is not initial.
        ls_even_footer-left_value = lo_ixml_hf_value_elem->get_value( ).
      endif.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'evenHeader' uri = namespace-main ).
      if lo_ixml_hf_value_elem is not initial.
        ls_even_header-left_value = lo_ixml_hf_value_elem->get_value( ).
      endif.

*        2do§1   Header/footer
      " TODO.. get the rest.

      io_worksheet->sheet_setup->set_header_footer( ip_odd_header   = ls_odd_header
                                                    ip_odd_footer   = ls_odd_footer
                                                    ip_even_header  = ls_even_header
                                                    ip_even_footer  = ls_even_footer ).

    endif.

    " Read pane
    lo_ixml_pane_elem = lo_ixml_sheetview_elem->find_from_name_ns( name = 'pane' uri = namespace-main ).
    if lo_ixml_pane_elem is bound.
      fill_struct_from_attributes( exporting ip_element = lo_ixml_pane_elem changing cp_structure = ls_excel_pane ).
      lv_pane_cell_col = ls_excel_pane-xsplit.
      lv_pane_cell_row = ls_excel_pane-ysplit.
      if    lv_pane_cell_col > 0
        and lv_pane_cell_row > 0.
        io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row
                                    ip_num_columns = lv_pane_cell_col ).
      elseif lv_pane_cell_row > 0.
        io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row ).
      else.
        io_worksheet->freeze_panes( ip_num_columns = lv_pane_cell_col ).
      endif.
      if ls_excel_pane-topleftcell is not initial.
        io_worksheet->set_pane_top_left_cell( ls_excel_pane-topleftcell ).
      endif.
    endif.

    " Start fix 276 Read data validations
    lo_ixml_datavalidations = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'dataValidation' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_datavalidations->create_iterator( ).
    lo_ixml_datavalidation_elem  ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_datavalidation_elem  is bound.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_datavalidation_elem
                                   changing
                                     cp_structure = ls_datavalidation ).
      clear lo_ixml_formula_elem.
      lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name_ns( name = 'formula1' uri = namespace-main ).
      if lo_ixml_formula_elem is bound.
        ls_datavalidation-formula1 = lo_ixml_formula_elem->get_value( ).
      endif.
      clear lo_ixml_formula_elem.
      lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name_ns( name = 'formula2' uri = namespace-main ).
      if lo_ixml_formula_elem is bound.
        ls_datavalidation-formula2 = lo_ixml_formula_elem->get_value( ).
      endif.
      split ls_datavalidation-sqref at space into table lt_datavalidation_range.
      loop at lt_datavalidation_range into lv_datavalidation_range.
        zcl_excel_common=>convert_range2column_a_row( exporting
                                                        i_range = lv_datavalidation_range
                                                      importing
                                                        e_column_start = ls_datavalidation-cell_column
                                                        e_column_end   = ls_datavalidation-cell_column_to
                                                        e_row_start    = ls_datavalidation-cell_row
                                                        e_row_end      = ls_datavalidation-cell_row_to ).
        lo_data_validation                   = io_worksheet->add_new_data_validation( ).
        lo_data_validation->type             = ls_datavalidation-type.
        lo_data_validation->allowblank       = ls_datavalidation-allowblank.
        if ls_datavalidation-showinputmessage is initial.
          lo_data_validation->showinputmessage = abap_false.
        else.
          lo_data_validation->showinputmessage = abap_true.
        endif.
        if ls_datavalidation-showerrormessage is initial.
          lo_data_validation->showerrormessage = abap_false.
        else.
          lo_data_validation->showerrormessage = abap_true.
        endif.
        if ls_datavalidation-showdropdown is initial.
          lo_data_validation->showdropdown = abap_false.
        else.
          lo_data_validation->showdropdown = abap_true.
        endif.
        lo_data_validation->operator         = ls_datavalidation-operator.
        lo_data_validation->formula1         = ls_datavalidation-formula1.
        lo_data_validation->formula2         = ls_datavalidation-formula2.
        lo_data_validation->prompttitle      = ls_datavalidation-prompttitle.
        lo_data_validation->prompt           = ls_datavalidation-prompt.
        lo_data_validation->errortitle       = ls_datavalidation-errortitle.
        lo_data_validation->error            = ls_datavalidation-error.
        lo_data_validation->errorstyle       = ls_datavalidation-errorstyle.
        lo_data_validation->cell_row         = ls_datavalidation-cell_row.
        lo_data_validation->cell_row_to      = ls_datavalidation-cell_row_to.
        lo_data_validation->cell_column      = ls_datavalidation-cell_column.
        lo_data_validation->cell_column_to   = ls_datavalidation-cell_column_to.
      endloop.
      lo_ixml_datavalidation_elem ?= lo_ixml_iterator->get_next( ).
    endwhile.
    " End fix 276 Read data validations

    " Read hyperlinks
    try.
        me->load_worksheet_hyperlinks( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet
                                       it_external_hyperlinks = lt_external_hyperlinks ).
      catch zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    endtry.

    try.
        me->fill_row_outlines( io_worksheet           = io_worksheet ).
      catch zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    endtry.

    " Issue #366 - conditional formatting
    try.
        me->load_worksheet_cond_format( io_ixml_worksheet      = lo_ixml_worksheet
                                        io_worksheet           = io_worksheet ).
      catch zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    endtry.

    " Issue #377 - pagebreaks
    try.
        me->load_worksheet_pagebreaks( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet ).
      catch zcx_excel. " Ignore pagebreak reading errors - pass everything we were able to identify
    endtry.

    try.
        me->load_worksheet_autofilter( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet ).
      catch zcx_excel. " Ignore autofilter reading errors - pass everything we were able to identify
    endtry.

    try.
        me->load_worksheet_ignored_errors( io_ixml_worksheet      = lo_ixml_worksheet
                                           io_worksheet           = io_worksheet ).
      catch zcx_excel. " Ignore "ignoredErrors" reading errors - pass everything we were able to identify
    endtry.

  endmethod.


  method load_worksheet_autofilter.

    types: begin of lty_autofilter,
             ref type string,
           end of lty_autofilter.

    data: lo_ixml_autofilter_elem    type ref to if_ixml_element,
          lv_ref                     type string,
          lo_ixml_filter_column_coll type ref to if_ixml_node_collection,
          lo_ixml_filter_column_iter type ref to if_ixml_node_iterator,
          lo_ixml_filter_column      type ref to if_ixml_element,
          lv_col_id                  type i,
          lv_column                  type zif_excel_data_decl=>zexcel_cell_column,
          lo_ixml_filters_coll       type ref to if_ixml_node_collection,
          lo_ixml_filters_iter       type ref to if_ixml_node_iterator,
          lo_ixml_filters            type ref to if_ixml_element,
          lo_ixml_filter_coll        type ref to if_ixml_node_collection,
          lo_ixml_filter_iter        type ref to if_ixml_node_iterator,
          lo_ixml_filter             type ref to if_ixml_element,
          lv_val                     type string,
          lo_autofilters             type ref to zcl_excel_autofilters,
          lo_autofilter              type ref to zcl_excel_autofilter.

    lo_autofilters = io_worksheet->excel->get_autofilters_reference( ).

    lo_ixml_autofilter_elem = io_ixml_worksheet->find_from_name_ns( name = 'autoFilter' uri = namespace-main ).
    if lo_ixml_autofilter_elem is bound.
      lv_ref = lo_ixml_autofilter_elem->get_attribute_ns( 'ref' ).

      lo_ixml_filter_column_coll = lo_ixml_autofilter_elem->get_elements_by_tag_name_ns( name = 'filterColumn' uri = namespace-main ).
      lo_ixml_filter_column_iter = lo_ixml_filter_column_coll->create_iterator( ).
      lo_ixml_filter_column ?= lo_ixml_filter_column_iter->get_next( ).
      while lo_ixml_filter_column is bound.
        lv_col_id = lo_ixml_filter_column->get_attribute_ns( 'colId' ).
        lv_column = lv_col_id + 1.

        lo_ixml_filters_coll = lo_ixml_filter_column->get_elements_by_tag_name_ns( name = 'filters' uri = namespace-main ).
        lo_ixml_filters_iter = lo_ixml_filters_coll->create_iterator( ).
        lo_ixml_filters ?= lo_ixml_filters_iter->get_next( ).
        while lo_ixml_filters is bound.

          lo_ixml_filter_coll = lo_ixml_filter_column->get_elements_by_tag_name_ns( name = 'filter' uri = namespace-main ).
          lo_ixml_filter_iter = lo_ixml_filter_coll->create_iterator( ).
          lo_ixml_filter ?= lo_ixml_filter_iter->get_next( ).
          while lo_ixml_filter is bound.
            lv_val = lo_ixml_filter->get_attribute_ns( 'val' ).

            lo_autofilter = lo_autofilters->get( io_worksheet = io_worksheet ).
            if lo_autofilter is not bound.
              lo_autofilter = lo_autofilters->add( io_sheet = io_worksheet ).
            endif.
            lo_autofilter->set_value(
                    i_column = lv_column
                    i_value  = lv_val ).

            lo_ixml_filter ?= lo_ixml_filter_iter->get_next( ).
          endwhile.

          lo_ixml_filters ?= lo_ixml_filters_iter->get_next( ).
        endwhile.

        lo_ixml_filter_column ?= lo_ixml_filter_column_iter->get_next( ).
      endwhile.
    endif.

  endmethod.


  method load_worksheet_cond_format.

    data: lo_ixml_cond_formats type ref to if_ixml_node_collection,
          lo_ixml_cond_format  type ref to if_ixml_element,
          lo_ixml_iterator     type ref to if_ixml_node_iterator,
          lo_ixml_rules        type ref to if_ixml_node_collection,
          lo_ixml_rule         type ref to if_ixml_element,
          lo_ixml_iterator2    type ref to if_ixml_node_iterator,
          lo_style_cond        type ref to zcl_excel_style_cond,
          lo_style_cond2       type ref to zcl_excel_style_cond.


    data: lv_area           type string,
          lt_areas          type standard table of string with non-unique default key,
          lv_area_start_row type zif_excel_data_decl=>zexcel_cell_row,
          lv_area_end_row   type zif_excel_data_decl=>zexcel_cell_row,
          lv_area_start_col type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_area_end_col   type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_rule           type zif_excel_data_decl=>zexcel_condition_rule.


    lo_ixml_cond_formats =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'conditionalFormatting' uri = namespace-main ).
    lo_ixml_iterator     =  lo_ixml_cond_formats->create_iterator( ).
    lo_ixml_cond_format  ?= lo_ixml_iterator->get_next( ).

    while lo_ixml_cond_format is bound.

      clear: lv_area,
             lo_ixml_rule,
             lo_style_cond.

*--------------------------------------------------------------------*
* Get type of rule
*--------------------------------------------------------------------*
      lo_ixml_rules       =  lo_ixml_cond_format->get_elements_by_tag_name_ns( name = 'cfRule' uri = namespace-main ).
      lo_ixml_iterator2   =  lo_ixml_rules->create_iterator( ).
      lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).

      while lo_ixml_rule is bound.
        lv_rule = lo_ixml_rule->get_attribute_ns( 'type' ).
        clear lo_style_cond.

*--------------------------------------------------------------------*
* Depending on ruletype get additional information
*--------------------------------------------------------------------*
        case lv_rule.

          when zcl_excel_style_cond=>c_rule_cellis.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_ci( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_databar.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_db( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_expression.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_ex( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_iconset.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_is( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_colorscale.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_cs( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_top10.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_t10( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          when zcl_excel_style_cond=>c_rule_above_average.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_aa(  io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).
          when others.
        endcase.

        if lo_style_cond is bound.
          lo_style_cond->rule      = lv_rule.
          lo_style_cond->priority  = lo_ixml_rule->get_attribute_ns( 'priority' ).
*--------------------------------------------------------------------*
* Set area to which conditional formatting belongs
*--------------------------------------------------------------------*
          lv_area =  lo_ixml_cond_format->get_attribute_ns( 'sqref' ).
          split lv_area at space into table lt_areas.
          delete lt_areas where table_line is initial.
          loop at lt_areas into lv_area.

            zcl_excel_common=>convert_range2column_a_row( exporting i_range        = lv_area
                                                          importing e_column_start = lv_area_start_col
                                                                    e_column_end   = lv_area_end_col
                                                                    e_row_start    = lv_area_start_row
                                                                    e_row_end      = lv_area_end_row   ).
            lo_style_cond->add_range( ip_start_column = lv_area_start_col
                                      ip_stop_column  = lv_area_end_col
                                      ip_start_row    = lv_area_start_row
                                      ip_stop_row     = lv_area_end_row   ).
          endloop.

        endif.
        lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).
      endwhile.


      lo_ixml_cond_format ?= lo_ixml_iterator->get_next( ).

    endwhile.

  endmethod.


  method load_worksheet_cond_format_aa.
    data: lv_dxf_style_index type i,
          val                type string.

    field-symbols: <ls_dxf_style> like line of me->mt_dxf_styles.

*--------------------------------------------------------------------*
* above or below average
*--------------------------------------------------------------------*
    val  = io_ixml_rule->get_attribute_ns( 'aboveAverage' ).
    if val = '0'.  " 0 = below average
      io_style_cond->mode_above_average-above_average = space.
    else.
      io_style_cond->mode_above_average-above_average = 'X'. " Not present or <> 0 --> we use above average
    endif.

*--------------------------------------------------------------------*
* Equal average also?
*--------------------------------------------------------------------*
    clear val.
    val  = io_ixml_rule->get_attribute_ns( 'equalAverage' ).
    if val = '1'.  " 0 = below average
      io_style_cond->mode_above_average-equal_average = 'X'.
    else.
      io_style_cond->mode_above_average-equal_average = ' '. " Not present or <> 1 --> we use not equal average
    endif.

*--------------------------------------------------------------------*
* Standard deviation instead of value ( 2nd stddev, 3rd stdev )
*--------------------------------------------------------------------*
    clear val.
    val  = io_ixml_rule->get_attribute_ns( 'stdDev' ).
    case val.
      when 1
        or 2
        or 3.  " These seem to be supported by excel - don't try anything more
        io_style_cond->mode_above_average-standard_deviation = val.
    endcase.

*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    read table me->mt_dxf_styles assigning <ls_dxf_style> with key dxf = lv_dxf_style_index.
    if sy-subrc = 0.
      io_style_cond->mode_above_average-cell_style = <ls_dxf_style>-guid.
    endif.

  endmethod.


  method load_worksheet_cond_format_ci.
    data: lo_ixml_nodes      type ref to if_ixml_node_collection,
          lo_ixml_iterator   type ref to if_ixml_node_iterator,
          lo_ixml            type ref to if_ixml_element,
          lv_dxf_style_index type i,
          lo_excel_style     like line of me->styles.

    field-symbols: <ls_dxf_style> like line of me->mt_dxf_styles.

    io_style_cond->mode_cellis-operator  = io_ixml_rule->get_attribute_ns( 'operator' ).
    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    read table me->mt_dxf_styles assigning <ls_dxf_style> with key dxf = lv_dxf_style_index.
    if sy-subrc = 0.
      io_style_cond->mode_cellis-cell_style = <ls_dxf_style>-guid.
    endif.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'formula' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_cellis-formula  = lo_ixml->get_value( ).

        when 2.
          io_style_cond->mode_cellis-formula2 = lo_ixml->get_value( ).

        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.


  endmethod.


  method load_worksheet_cond_format_cs.
    data: lo_ixml_nodes    type ref to if_ixml_node_collection,
          lo_ixml_iterator type ref to if_ixml_node_iterator,
          lo_ixml          type ref to if_ixml_element.


    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'cfvo' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_colorscale-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_colorscale-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

        when 2.
          io_style_cond->mode_colorscale-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_colorscale-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        when 3.
          io_style_cond->mode_colorscale-cfvo3_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_colorscale-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'color' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_colorscale-colorrgb1  = lo_ixml->get_attribute_ns( 'rgb' ).

        when 2.
          io_style_cond->mode_colorscale-colorrgb2  = lo_ixml->get_attribute_ns( 'rgb' ).

        when 3.
          io_style_cond->mode_colorscale-colorrgb3  = lo_ixml->get_attribute_ns( 'rgb' ).

        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.

  endmethod.


  method load_worksheet_cond_format_db.
    data: lo_ixml_nodes    type ref to if_ixml_node_collection,
          lo_ixml_iterator type ref to if_ixml_node_iterator,
          lo_ixml          type ref to if_ixml_element.

    lo_ixml ?= io_ixml_rule->find_from_name_ns( name = 'color' uri = namespace-main ).
    if lo_ixml is bound.
      io_style_cond->mode_databar-colorrgb = lo_ixml->get_attribute_ns( 'rgb' ).
    endif.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'cfvo' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_databar-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_databar-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

        when 2.
          io_style_cond->mode_databar-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_databar-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.


  endmethod.


  method load_worksheet_cond_format_ex.
    data: lo_ixml_nodes      type ref to if_ixml_node_collection,
          lo_ixml_iterator   type ref to if_ixml_node_iterator,
          lo_ixml            type ref to if_ixml_element,
          lv_dxf_style_index type i,
          lo_excel_style     like line of me->styles.

    field-symbols: <ls_dxf_style> like line of me->mt_dxf_styles.

    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    read table me->mt_dxf_styles assigning <ls_dxf_style> with key dxf = lv_dxf_style_index.
    if sy-subrc = 0.
      io_style_cond->mode_expression-cell_style = <ls_dxf_style>-guid.
    endif.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'formula' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_expression-formula  = lo_ixml->get_value( ).


        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.


  endmethod.


  method load_worksheet_cond_format_is.
    data: lo_ixml_nodes        type ref to if_ixml_node_collection,
          lo_ixml_iterator     type ref to if_ixml_node_iterator,
          lo_ixml              type ref to if_ixml_element,
          lo_ixml_rule_iconset type ref to if_ixml_element.

    lo_ixml_rule_iconset ?= io_ixml_rule->get_first_child( ).
    io_style_cond->mode_iconset-iconset   = lo_ixml_rule_iconset->get_attribute_ns( 'iconSet' ).
    io_style_cond->mode_iconset-showvalue = lo_ixml_rule_iconset->get_attribute_ns( 'showValue' ).
    lo_ixml_nodes ?= lo_ixml_rule_iconset->get_elements_by_tag_name_ns( name = 'cfvo' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    while lo_ixml is bound.

      case sy-index.
        when 1.
          io_style_cond->mode_iconset-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

        when 2.
          io_style_cond->mode_iconset-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        when 3.
          io_style_cond->mode_iconset-cfvo3_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo3_value = lo_ixml->get_attribute_ns( 'val' ).

        when 4.
          io_style_cond->mode_iconset-cfvo4_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo4_value = lo_ixml->get_attribute_ns( 'val' ).

        when 5.
          io_style_cond->mode_iconset-cfvo5_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo5_value = lo_ixml->get_attribute_ns( 'val' ).

        when others.
          exit.
      endcase.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    endwhile.

  endmethod.


  method load_worksheet_cond_format_t10.
    data: lv_dxf_style_index type i.

    field-symbols: <ls_dxf_style> like line of me->mt_dxf_styles.

    io_style_cond->mode_top10-topxx_count  = io_ixml_rule->get_attribute_ns( 'rank' ).        " Top10, Top20, Top 50...

    io_style_cond->mode_top10-percent      = io_ixml_rule->get_attribute_ns( 'percent' ).     " Top10 percent instead of Top10 values
    if io_style_cond->mode_top10-percent = '1'.
      io_style_cond->mode_top10-percent = 'X'.
    else.
      io_style_cond->mode_top10-percent = ' '.
    endif.

    io_style_cond->mode_top10-bottom       = io_ixml_rule->get_attribute_ns( 'bottom' ).      " Bottom10 instead of Top10
    if io_style_cond->mode_top10-bottom = '1'.
      io_style_cond->mode_top10-bottom = 'X'.
    else.
      io_style_cond->mode_top10-bottom = ' '.
    endif.
*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    read table me->mt_dxf_styles assigning <ls_dxf_style> with key dxf = lv_dxf_style_index.
    if sy-subrc = 0.
      io_style_cond->mode_top10-cell_style = <ls_dxf_style>-guid.
    endif.

  endmethod.


  method load_worksheet_drawing.

*    TYPES: BEGIN OF t_c_nv_pr,
*             name TYPE string,
*             id   TYPE string,
*           END OF t_c_nv_pr.
*
*    TYPES: BEGIN OF t_blip,
*             cstate TYPE string,
*             embed  TYPE string,
*           END OF t_blip.
*
*    TYPES: BEGIN OF t_chart,
*             id TYPE string,
*           END OF t_chart.
*
*    CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
*               lc_xml_attr_true_int TYPE string VALUE '1'.
*    CONSTANTS: lc_rel_chart TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
*               lc_rel_image TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.
*
*    DATA: drawing           TYPE REF TO if_ixml_document,
*          anchors           TYPE REF TO if_ixml_node_collection,
*          node              TYPE REF TO if_ixml_element,
*          coll_length       TYPE i,
*          iterator          TYPE REF TO if_ixml_node_iterator,
*          anchor_elem       TYPE REF TO if_ixml_element,
*
*          relationship      TYPE t_relationship,
*          rel_drawings      TYPE t_rel_drawings,
*          rel_drawing       TYPE t_rel_drawing,
*          rels_drawing      TYPE REF TO if_ixml_document,
*          rels_drawing_path TYPE string,
*          stripped_name     TYPE string,
*          dirname           TYPE string,
*
*          path              TYPE string,
*          path2             TYPE text255,
*          file_ext2         TYPE char10.
*
*    " Read Workbook Relationships
*    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
*      EXPORTING
*        full_name     = ip_path
*      IMPORTING
*        stripped_name = stripped_name
*        file_path     = dirname.
*    CONCATENATE dirname '_rels/' stripped_name '.rels'
*      INTO rels_drawing_path.
*    rels_drawing_path = resolve_path( rels_drawing_path ).
*    rels_drawing = me->get_ixml_from_zip_archive( rels_drawing_path ).
*    node ?= rels_drawing->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
*    WHILE node IS BOUND.
*      fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = relationship ).
*
*      rel_drawing-id = relationship-id.
*
*      CONCATENATE dirname relationship-target INTO path.
*      path = resolve_path( path ).
*      rel_drawing-content = me->get_from_zip_archive( path ). "------------> This is for template usage
*
*      path2 = path.
*      zcl_excel_common=>split_file( EXPORTING ip_file = path2
*                                    IMPORTING ep_extension = file_ext2 ).
*      rel_drawing-file_ext = file_ext2.
*
*      "-------------Added by Alessandro Iannacci - Should load graph xml
*      CASE relationship-type.
*        WHEN lc_rel_chart.
*          "Read chart xml
*          rel_drawing-content_xml = me->get_ixml_from_zip_archive( path ).
*        WHEN OTHERS.
*      ENDCASE.
*      "----------------------------
*
*
*      APPEND rel_drawing TO rel_drawings.
*
*      node ?= node->get_next( ).
*    ENDWHILE.
*
*    drawing = me->get_ixml_from_zip_archive( ip_path ).
*
** one-cell anchor **************
*    anchors = drawing->get_elements_by_tag_name_ns( name = 'oneCellAnchor' uri = namespace-xdr ).
*    coll_length = anchors->get_length( ).
*    iterator = anchors->create_iterator( ).
*    DO coll_length TIMES.
*      anchor_elem ?= iterator->get_next( ).
*
*      CALL METHOD me->load_drawing_anchor
*        EXPORTING
*          io_anchor_element   = anchor_elem
*          io_worksheet        = io_worksheet
*          it_related_drawings = rel_drawings.
*
*    ENDDO.
*
** two-cell anchor ******************
*    anchors = drawing->get_elements_by_tag_name_ns( name = 'twoCellAnchor' uri = namespace-xdr ).
*    coll_length = anchors->get_length( ).
*    iterator = anchors->create_iterator( ).
*    DO coll_length TIMES.
*      anchor_elem ?= iterator->get_next( ).
*
*      CALL METHOD me->load_drawing_anchor
*        EXPORTING
*          io_anchor_element   = anchor_elem
*          io_worksheet        = io_worksheet
*          it_related_drawings = rel_drawings.
*
*    ENDDO.

  endmethod.

  method load_comments.
    data: lo_comments_xml       type ref to if_ixml_document,
          lo_node_comment       type ref to if_ixml_element,
          lo_node_comment_child type ref to if_ixml_element,
          lo_node_r_child_t     type ref to if_ixml_element,
          lo_attr               type ref to if_ixml_attribute,
          lo_comment            type ref to zcl_excel_comment,
          lv_comment_text       type string,
          lv_node_value         type string,
          lv_attr_value         type string.

    lo_comments_xml = me->get_ixml_from_zip_archive( ip_path ).

    lo_node_comment ?= lo_comments_xml->find_from_name_ns( name = 'comment' uri = namespace-main ).
    while lo_node_comment is bound.

      clear lv_comment_text.
      lo_attr = lo_node_comment->get_attribute_node_ns( name = 'ref' ).
      lv_attr_value  = lo_attr->get_value( ).

      lo_node_comment_child ?= lo_node_comment->get_first_child( ).
      while lo_node_comment_child is bound.
        " There will be rPr nodes here, but we do not support them
        " in comments right now; see 'load_shared_strings' for handling.
        " Extract the <t>...</t> part of each <r>-tag
        lo_node_r_child_t ?= lo_node_comment_child->find_from_name_ns( name = 't' uri = namespace-main ).
        if lo_node_r_child_t is bound.
          lv_node_value = lo_node_r_child_t->get_value( ).
          concatenate lv_comment_text lv_node_value into lv_comment_text respecting blanks.
        endif.
        lo_node_comment_child ?= lo_node_comment_child->get_next( ).
      endwhile.

      create object lo_comment.
      lo_comment->set_text( ip_ref = lv_attr_value ip_text = lv_comment_text ).
      io_worksheet->add_comment( lo_comment ).

      lo_node_comment ?= lo_node_comment->get_next( ).
    endwhile.

  endmethod.

  method load_worksheet_hyperlinks.

    data: lo_ixml_hyperlinks type ref to if_ixml_node_collection,
          lo_ixml_hyperlink  type ref to if_ixml_element,
          lo_ixml_iterator   type ref to if_ixml_node_iterator,
          lv_row_start       type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_end         type zif_excel_data_decl=>zexcel_cell_row,
          lv_column_start    type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_column_end      type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_is_internal     type abap_bool,
          lv_url             type string,
          lv_value           type zif_excel_data_decl=>zexcel_cell_value.

    data: begin of ls_hyperlink,
            ref      type string,
            display  type string,
            location type string,
            tooltip  type string,
            r_id     type string,
          end of ls_hyperlink.

    field-symbols: <ls_external_hyperlink> like line of it_external_hyperlinks.

    lo_ixml_hyperlinks =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'hyperlink' uri = namespace-main ).
    lo_ixml_iterator   =  lo_ixml_hyperlinks->create_iterator( ).
    lo_ixml_hyperlink  ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_hyperlink is bound.

      clear ls_hyperlink.
      clear lv_url.

      ls_hyperlink-ref      = lo_ixml_hyperlink->get_attribute_ns( 'ref' ).
      ls_hyperlink-display  = lo_ixml_hyperlink->get_attribute_ns( 'display' ).
      ls_hyperlink-location = lo_ixml_hyperlink->get_attribute_ns( 'location' ).
      ls_hyperlink-tooltip  = lo_ixml_hyperlink->get_attribute_ns( 'tooltip' ).
      ls_hyperlink-r_id     = lo_ixml_hyperlink->get_attribute_ns( name = 'id' uri = namespace-r ).
      if ls_hyperlink-r_id is initial.  " Internal link
        lv_is_internal = abap_true.
        lv_url = ls_hyperlink-location.
      else.                             " External link
        read table it_external_hyperlinks assigning <ls_external_hyperlink> with table key id = ls_hyperlink-r_id.
        if sy-subrc = 0.
          lv_is_internal = abap_false.
          lv_url = <ls_external_hyperlink>-target.
        endif.
      endif.

      if lv_url is not initial.  " because of unsupported external links

        zcl_excel_common=>convert_range2column_a_row(
          exporting
            i_range        = ls_hyperlink-ref
          importing
            e_column_start = lv_column_start
            e_column_end   = lv_column_end
            e_row_start    = lv_row_start
            e_row_end      = lv_row_end ).

        io_worksheet->set_area_hyperlink(
          exporting
            ip_column_start = lv_column_start
            ip_column_end   = lv_column_end
            ip_row          = lv_row_start
            ip_row_to       = lv_row_end
            ip_url          = lv_url
            ip_is_internal  = lv_is_internal ).

      endif.

      lo_ixml_hyperlink ?= lo_ixml_iterator->get_next( ).

    endwhile.


  endmethod.


  method load_worksheet_ignored_errors.

    data: lo_ixml_ignored_errors type ref to if_ixml_node_collection,
          lo_ixml_ignored_error  type ref to if_ixml_element,
          lo_ixml_iterator       type ref to if_ixml_node_iterator,
          ls_ignored_error       type zcl_excel_worksheet=>mty_s_ignored_errors,
          lt_ignored_errors      type zcl_excel_worksheet=>mty_th_ignored_errors.

    data: begin of ls_raw_ignored_error,
            sqref              type string,
            evalerror          type string,
            twodigittextyear   type string,
            numberstoredastext type string,
            formula            type string,
            formularange       type string,
            unlockedformula    type string,
            emptycellreference type string,
            listdatavalidation type string,
            calculatedcolumn   type string,
          end of ls_raw_ignored_error.

    clear lt_ignored_errors.

    lo_ixml_ignored_errors =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'ignoredError' uri = namespace-main ).
    lo_ixml_iterator   =  lo_ixml_ignored_errors->create_iterator( ).
    lo_ixml_ignored_error  ?= lo_ixml_iterator->get_next( ).

    while lo_ixml_ignored_error is bound.

      fill_struct_from_attributes( exporting
                                     ip_element   = lo_ixml_ignored_error
                                   changing
                                     cp_structure = ls_raw_ignored_error ).

      clear ls_ignored_error.
      ls_ignored_error-cell_coords = ls_raw_ignored_error-sqref.
      ls_ignored_error-eval_error = boolc( ls_raw_ignored_error-evalerror = '1' ).
      ls_ignored_error-two_digit_text_year = boolc( ls_raw_ignored_error-twodigittextyear = '1' ).
      ls_ignored_error-number_stored_as_text = boolc( ls_raw_ignored_error-numberstoredastext = '1' ).
      ls_ignored_error-formula = boolc( ls_raw_ignored_error-formula = '1' ).
      ls_ignored_error-formula_range = boolc( ls_raw_ignored_error-formularange = '1' ).
      ls_ignored_error-unlocked_formula = boolc( ls_raw_ignored_error-unlockedformula = '1' ).
      ls_ignored_error-empty_cell_reference = boolc( ls_raw_ignored_error-emptycellreference = '1' ).
      ls_ignored_error-list_data_validation = boolc( ls_raw_ignored_error-listdatavalidation = '1' ).
      ls_ignored_error-calculated_column  = boolc( ls_raw_ignored_error-calculatedcolumn = '1' ).

      insert ls_ignored_error into table lt_ignored_errors.

      lo_ixml_ignored_error ?= lo_ixml_iterator->get_next( ).

    endwhile.

    io_worksheet->set_ignored_errors( lt_ignored_errors ).

  endmethod.


  method load_worksheet_pagebreaks.

    data: lo_node           type ref to if_ixml_element,
          lo_ixml_rowbreaks type ref to if_ixml_node_collection,
          lo_ixml_colbreaks type ref to if_ixml_node_collection,
          lo_ixml_iterator  type ref to if_ixml_node_iterator,
          lo_ixml_rowbreak  type ref to if_ixml_element,
          lo_ixml_colbreak  type ref to if_ixml_element,
          lo_style_cond     type ref to zcl_excel_style_cond,
          lv_count          type i.


    data: lt_pagebreaks type standard table of zcl_excel_worksheet_pagebreaks=>ts_pagebreak_at,
          lo_pagebreaks type ref to zcl_excel_worksheet_pagebreaks.

    field-symbols: <ls_pagebreak_row> like line of lt_pagebreaks.
    field-symbols: <ls_pagebreak_col> like line of lt_pagebreaks.

*--------------------------------------------------------------------*
* Get minimal number of cells where to add pagebreaks
* Since rows and columns are handled in separate nodes
* Build table to identify these cells
*--------------------------------------------------------------------*
    lo_node ?= io_ixml_worksheet->find_from_name_ns( name = 'rowBreaks' uri = namespace-main ).
    check lo_node is bound.
    lo_ixml_rowbreaks =  lo_node->get_elements_by_tag_name_ns( name = 'brk' uri = namespace-main ).
    lo_ixml_iterator  =  lo_ixml_rowbreaks->create_iterator( ).
    lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
    while lo_ixml_rowbreak is bound.
      append initial line to lt_pagebreaks assigning <ls_pagebreak_row>.
      <ls_pagebreak_row>-cell_row = lo_ixml_rowbreak->get_attribute_ns( 'id' ).

      lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
    endwhile.
    check <ls_pagebreak_row> is assigned.

    lo_node ?= io_ixml_worksheet->find_from_name_ns( name = 'colBreaks' uri = namespace-main ).
    check lo_node is bound.
    lo_ixml_colbreaks =  lo_node->get_elements_by_tag_name_ns( name = 'brk' uri = namespace-main ).
    lo_ixml_iterator  =  lo_ixml_colbreaks->create_iterator( ).
    lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
    clear lv_count.
    while lo_ixml_colbreak is bound.
      lv_count += 1.
      read table lt_pagebreaks index lv_count assigning <ls_pagebreak_col>.
      if sy-subrc <> 0.
        append initial line to lt_pagebreaks assigning <ls_pagebreak_col>.
        <ls_pagebreak_col>-cell_row = <ls_pagebreak_row>-cell_row.
      endif.
      <ls_pagebreak_col>-cell_column = lo_ixml_colbreak->get_attribute_ns( 'id' ).

      lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
    endwhile.
*--------------------------------------------------------------------*
* Finally add each pagebreak
*--------------------------------------------------------------------*
    lo_pagebreaks = io_worksheet->get_pagebreaks( ).
    loop at lt_pagebreaks assigning <ls_pagebreak_row>.
      lo_pagebreaks->add_pagebreak( ip_column = <ls_pagebreak_row>-cell_column
                                    ip_row    = <ls_pagebreak_row>-cell_row ).
    endloop.


  endmethod.


  method load_worksheet_pagemargins.

    types: begin of lty_page_margins,
             footer type string,
             header type string,
             bottom type string,
             top    type string,
             right  type string,
             left   type string,
           end of lty_page_margins.

    data:lo_ixml_pagemargins_elem type ref to if_ixml_element,
         ls_pagemargins           type lty_page_margins.


    lo_ixml_pagemargins_elem = io_ixml_worksheet->find_from_name_ns( name = 'pageMargins' uri = namespace-main ).
    if lo_ixml_pagemargins_elem is not initial.
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_pagemargins_elem
                                   changing
                                     cp_structure = ls_pagemargins ).
      io_worksheet->sheet_setup->margin_bottom = zcl_excel_common=>excel_string_to_number( ls_pagemargins-bottom ).
      io_worksheet->sheet_setup->margin_footer = zcl_excel_common=>excel_string_to_number( ls_pagemargins-footer ).
      io_worksheet->sheet_setup->margin_header = zcl_excel_common=>excel_string_to_number( ls_pagemargins-header ).
      io_worksheet->sheet_setup->margin_left   = zcl_excel_common=>excel_string_to_number( ls_pagemargins-left   ).
      io_worksheet->sheet_setup->margin_right  = zcl_excel_common=>excel_string_to_number( ls_pagemargins-right  ).
      io_worksheet->sheet_setup->margin_top    = zcl_excel_common=>excel_string_to_number( ls_pagemargins-top    ).
    endif.

  endmethod.


  method load_worksheet_tables.

    data lo_ixml_table_columns type ref to if_ixml_node_collection.
    data lo_ixml_table_column  type ref to if_ixml_element.
    data lo_ixml_table type ref to if_ixml_element.
    data lo_ixml_table_style type ref to if_ixml_element.
    data lt_field_catalog type zif_excel_data_decl=>zexcel_t_fieldcatalog.
    data ls_field_catalog type zif_excel_data_decl=>zexcel_s_fieldcatalog.
    data lo_ixml_iterator type ref to if_ixml_node_iterator.
    data ls_table_settings type zif_excel_data_decl=>zexcel_s_table_settings.
    data lv_path type string.
    data lt_components type abap_component_tab.
    data ls_component type abap_componentdescr.
    data lo_rtti_table type ref to cl_abap_tabledescr.
    data lv_dref_table type ref to data.
    data lv_num_lines type i.
    data lo_line_type type ref to cl_abap_structdescr.

    data: begin of ls_table,
            id             type string,
            name           type string,
            displayname    type string,
            ref            type string,
            totalsrowshown type string,
          end of ls_table.

    data: begin of ls_table_style,
            name              type string,
            showrowstripes    type string,
            showcolumnstripes type string,
          end of ls_table_style.

    data: begin of ls_table_column,
            id   type string,
            name type string,
          end of ls_table_column.

    field-symbols <ls_table> like line of it_tables.
    field-symbols <lt_table> type standard table.
    field-symbols <ls_field> type zif_excel_data_decl=>zexcel_s_fieldcatalog.

    loop at it_tables assigning <ls_table>.

      concatenate iv_dirname <ls_table>-target into lv_path.
      lv_path = resolve_path( lv_path ).

      lo_ixml_table = me->get_ixml_from_zip_archive( lv_path )->get_root_element( ).
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_table
                                   changing
                                     cp_structure = ls_table ).

      lo_ixml_table_style ?= lo_ixml_table->find_from_name( 'tableStyleInfo' ).
      fill_struct_from_attributes( exporting
                                     ip_element = lo_ixml_table_style
                                   changing
                                     cp_structure = ls_table_style ).

      ls_table_settings-table_name = ls_table-name.
      ls_table_settings-table_style = ls_table_style-name.
      ls_table_settings-show_column_stripes = boolc( ls_table_style-showcolumnstripes = '1' ).
      ls_table_settings-show_row_stripes = boolc( ls_table_style-showrowstripes = '1' ).

      zcl_excel_common=>convert_range2column_a_row(
        exporting
          i_range        = ls_table-ref
        importing
          e_column_start = ls_table_settings-top_left_column
          e_column_end   = ls_table_settings-bottom_right_column
          e_row_start    = ls_table_settings-top_left_row
          e_row_end      = ls_table_settings-bottom_right_row ).

      lo_ixml_table_columns =  lo_ixml_table->get_elements_by_tag_name( name = 'tableColumn' ).
      lo_ixml_iterator     =  lo_ixml_table_columns->create_iterator( ).
      lo_ixml_table_column  ?= lo_ixml_iterator->get_next( ).
      clear lt_field_catalog.
      while lo_ixml_table_column is bound.

        clear ls_table_column.
        fill_struct_from_attributes( exporting
                                       ip_element = lo_ixml_table_column
                                     changing
                                       cp_structure = ls_table_column ).

        ls_field_catalog-position = lines( lt_field_catalog ) + 1.
        ls_field_catalog-fieldname = |COMP_{ ls_field_catalog-position pad = '0' align = right width = 4 }|.
        ls_field_catalog-scrtext_l = ls_table_column-name.
        ls_field_catalog-abap_type = cl_abap_typedescr=>typekind_string.
        append ls_field_catalog to lt_field_catalog.

        lo_ixml_table_column ?= lo_ixml_iterator->get_next( ).

      endwhile.

      clear lt_components.
      loop at lt_field_catalog assigning <ls_field>.
        clear ls_component.
        ls_component-name = <ls_field>-fieldname.
        ls_component-type = cl_abap_elemdescr=>get_string( ).
        append ls_component to lt_components.
      endloop.

      lo_line_type = cl_abap_structdescr=>get( lt_components ).
      lo_rtti_table = cl_abap_tabledescr=>get( lo_line_type ).
      create data lv_dref_table type handle lo_rtti_table.
      assign lv_dref_table->* to <lt_table>.

      lv_num_lines = ls_table_settings-bottom_right_row - ls_table_settings-top_left_row.
      do lv_num_lines times.
        append initial line to <lt_table>.
      enddo.

      io_worksheet->bind_table(
        exporting
          ip_table            = <lt_table>
          it_field_catalog    = lt_field_catalog
          is_table_settings   = ls_table_settings ).

    endloop.

  endmethod.


  method read_from_applserver.
    raise exception type zcx_excel
      exporting
        error = 'Not supported in Cloud'.
  endmethod.


  method read_from_local_file.
    raise exception type zcx_excel
      exporting
        error = 'Not supported in Cloud'.
  endmethod.


  method resolve_path.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Determine whether the replacement should be done
*                iterative to allow /../../..   or something alike
*        2do§2   Determine whether /./ has to be supported as well
*        2do§3   Create unit-test for this method
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: replaced previous coding by regular expression
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1  This routine will receive a path, that may have a relative pathname (/../) included somewhere
*     The output should be a resolved path without relative references
*     Example:  Input     xl/worksheets/../drawings/drawing1.xml
*               Output    xl/drawings/drawing1.xml
*--------------------------------------------------------------------*

    rp_result = ip_path.
*--------------------------------------------------------------------*
* §1  Remove relative pathnames
*--------------------------------------------------------------------*
*  Regular expression   [^/]*/\.\./
*                       [^/]*            --> any number of characters other than /
*   followed by              /\.\./      --> the sequence /../
*   ==> worksheets/../ will be found in the example
*--------------------------------------------------------------------*
    replace PCRE  '[^/]*/\.\./' in rp_result with ``.


  endmethod.


  method resolve_referenced_formulae.
    types: begin of ty_referenced_cells,
             sheet    type ref to zcl_excel_worksheet,
             si       type i,
             row_from type i,
             row_to   type i,
             col_from type i,
             col_to   type i,
             formula  type string,
             ref_cell type c length 10,
           end of ty_referenced_cells.

    data: ls_ref_formula       like line of me->mt_ref_formulae,
          lts_referenced_cells type sorted table of ty_referenced_cells with non-unique key sheet si row_from row_to col_from col_to,
          ls_referenced_cell   like line of lts_referenced_cells,
          lv_col_from          type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_col_to            type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_resulting_formula type string,
          lv_current_cell      type c length 10.


    me->mt_ref_formulae = me->mt_ref_formulae.

*--------------------------------------------------------------------*
* Get referenced Cells,  Build ranges for easy lookup
*--------------------------------------------------------------------*
    loop at me->mt_ref_formulae into ls_ref_formula where ref <> space. "#EC CI_HASHSEQ

      clear ls_referenced_cell.
      ls_referenced_cell-sheet      = ls_ref_formula-sheet.
      ls_referenced_cell-si         = ls_ref_formula-si.
      ls_referenced_cell-formula    = ls_ref_formula-formula.

      try.
          zcl_excel_common=>convert_range2column_a_row( exporting i_range        = ls_ref_formula-ref
                                                        importing e_column_start = lv_col_from
                                                                  e_column_end   = lv_col_to
                                                                  e_row_start    = ls_referenced_cell-row_from
                                                                  e_row_end      = ls_referenced_cell-row_to  ).
          ls_referenced_cell-col_from = zcl_excel_common=>convert_column2int( lv_col_from ).
          ls_referenced_cell-col_to   = zcl_excel_common=>convert_column2int( lv_col_to ).


          clear ls_referenced_cell-ref_cell.
          try.
              ls_referenced_cell-ref_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
              ls_referenced_cell-ref_cell+3  = ls_ref_formula-row.
              condense ls_referenced_cell-ref_cell no-gaps.
            catch zcx_excel.
          endtry.

          insert ls_referenced_cell into table lts_referenced_cells.
        catch zcx_excel.
      endtry.

    endloop.

*  break x0009004.
*--------------------------------------------------------------------*
* For each referencing cell determine the referenced cell
* and resolve the formula
*--------------------------------------------------------------------*
    loop at me->mt_ref_formulae into ls_ref_formula where ref = space. "#EC CI_HASHSEQ


      clear lv_current_cell.
      try.
          lv_current_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
          lv_current_cell+3  = ls_ref_formula-row.
          condense lv_current_cell no-gaps.
        catch zcx_excel.
      endtry.

      loop at lts_referenced_cells into ls_referenced_cell where sheet     = ls_ref_formula-sheet
                                                             and si        = ls_ref_formula-si
                                                             and row_from <= ls_ref_formula-row
                                                             and row_to   >= ls_ref_formula-row
                                                             and col_from <= ls_ref_formula-column
                                                             and col_to   >= ls_ref_formula-column.

        try.

            lv_resulting_formula = zcl_excel_common=>determine_resulting_formula( iv_reference_cell     = ls_referenced_cell-ref_cell
                                                                                  iv_reference_formula  = ls_referenced_cell-formula
                                                                                  iv_current_cell       = lv_current_cell ).

            ls_referenced_cell-sheet->set_cell_formula( ip_column   = ls_ref_formula-column
                                                        ip_row      = ls_ref_formula-row
                                                        ip_formula  = lv_resulting_formula ).
          catch zcx_excel.
        endtry.
        exit.

      endloop.

    endloop.
  endmethod.


  method unescape_string_value.

    data:
      "Marks the Position before the searched Pattern occurs in the String
      "For example in String A_X_TEST_X, the Table is filled with 1 and 8
      lt_character_positions       type table of i,
      lv_character_position        type i,
      lv_character_position_plus_2 type i,
      lv_character_position_plus_6 type i,
      lv_unescaped_value           type string.

    " The text "_x...._", with "_x" not "_X". Each "." represents one character, being 0-9 a-f or A-F (case insensitive),
    " is interpreted like Unicode character U+.... (e.g. "_x0041_" is rendered like "A") is for characters.
    " To not interpret it, Excel replaces the first "_" with "_x005f_".
    result = i_value.

    if provided_string_is_escaped( i_value ) = abap_true.
      clear lt_character_positions.
      append sy-fdpos to lt_character_positions.
      lv_character_position = sy-fdpos + 1.
      while result+lv_character_position cs '_x'.
        lv_character_position += sy-fdpos.
        append lv_character_position to lt_character_positions.
        lv_character_position += 1.
      endwhile.
      sort lt_character_positions by table_line descending.
      loop at lt_character_positions into lv_character_position.
        lv_character_position_plus_2 = lv_character_position + 2.
        lv_character_position_plus_6 = lv_character_position + 6.
        if substring( val = result off = lv_character_position_plus_2 len = 4 ) co '0123456789ABCDEFabcdef'.
          if substring( val = result off = lv_character_position_plus_6 len = 1 ) = '_'.
            lv_unescaped_value = cl_abap_conv_codepage=>create_out( codepage = `UTF-8`
                        )->convert( source = to_upper( substring( val = result off = lv_character_position_plus_2 len = 4 ) ) ).
            replace section offset lv_character_position length 7 of result with lv_unescaped_value.
          endif.
        endif.
      endloop.
    endif.

  endmethod.


  method zif_excel_reader~load.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Map Document Properties to ZCL_EXCEL
*--------------------------------------------------------------------*

    constants: lcv_core_properties type string value 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
               lcv_office_document type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'.

    data: lo_rels         type ref to if_ixml_document,
          lo_node         type ref to if_ixml_element,
          ls_relationship type t_relationship.

*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller

* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets

* §3  Extracting from this the path&file where the workbook is located
*     Following is an example how this file could be set up
*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*            <Relationship Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Id="rId2"/>
*            <Relationship Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/>
*        </Relationships>
*--------------------------------------------------------------------*

    clear mt_dxf_styles.
    clear mt_ref_formulae.
    clear shared_strings.
    clear styles.

*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller
*--------------------------------------------------------------------*
    if iv_zcl_excel_classname is initial.
      create object r_excel.
    else.
      create object r_excel type (iv_zcl_excel_classname).
    endif.

    zip = create_zip_archive( i_xlsx_binary = i_excel2007
                              i_use_alternate_zip = i_use_alternate_zip ).

*--------------------------------------------------------------------*
* §2  Get file in folderstructure
*--------------------------------------------------------------------*
    lo_rels = get_ixml_from_zip_archive( '_rels/.rels' ).

*--------------------------------------------------------------------*
* §3  Cycle through the Relationship Tags and use the ones we need
*--------------------------------------------------------------------*
    lo_node ?= lo_rels->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ). "#EC NOTEXT
    while lo_node is bound.

      fill_struct_from_attributes( exporting
                                     ip_element   = lo_node
                                   changing
                                     cp_structure = ls_relationship ).
      case ls_relationship-type.

        when lcv_office_document.
*--------------------------------------------------------------------*
* Parse workbook - main part here
*--------------------------------------------------------------------*
          load_workbook( iv_workbook_full_filename  = ls_relationship-target
                         io_excel                   = r_excel ).

        when lcv_core_properties.
          " 2do§1   Map Document Properties to ZCL_EXCEL

        when others.

      endcase.
      lo_node ?= lo_node->get_next( ).

    endwhile.


  endmethod.


  method zif_excel_reader~load_file.

    data: lv_excel_data type xstring.

*--------------------------------------------------------------------*
* Read file into binary string
*--------------------------------------------------------------------*
    if i_from_applserver = abap_true.
      lv_excel_data = read_from_applserver( i_filename ).
    else.
      lv_excel_data = read_from_local_file( i_filename ).
    endif.

*--------------------------------------------------------------------*
* Parse Excel data into ZCL_EXCEL object from binary string
*--------------------------------------------------------------------*
    r_excel = zif_excel_reader~load( i_excel2007            = lv_excel_data
                                     i_use_alternate_zip    = i_use_alternate_zip
                                     iv_zcl_excel_classname = iv_zcl_excel_classname ).

  endmethod.
  method provided_string_is_escaped.

    "Check if passed value is really an escaped Character
    if value cs '_x'.
      is_escaped = abap_true.
      try.
          if substring( val = value off = sy-fdpos + 6 len = 1 ) <> '_'.
            is_escaped = abap_false.
          endif.
        catch cx_sy_range_out_of_bounds.
          is_escaped = abap_false.
      endtry.
    endif.
  endmethod.

endclass.
