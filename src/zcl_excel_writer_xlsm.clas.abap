class zcl_excel_writer_xlsm definition
  public
  inheriting from zcl_excel_writer_2007
  create public .

  public section.
*"* public components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
  protected section.

*"* protected components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
    constants c_xl_vbaproject type string value 'xl/vbaProject.bin'. "#EC NOTEXT

    methods add_further_data_to_zip
        redefinition .
    methods create
        redefinition .
    methods create_content_types
        redefinition .
    methods create_xl_relationships
        redefinition .
    methods create_xl_sheet
        redefinition .
    methods create_xl_workbook
        redefinition .
  private section.
*"* private components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
endclass.



class zcl_excel_writer_xlsm implementation.


  method add_further_data_to_zip.

    super->add_further_data_to_zip( io_zip = io_zip ).

* Add vbaProject.bin to zip
    io_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

  endmethod.


  method create.


* Office 2007 file format is a cab of several xml files with extension .xlsx

    data: lo_zip              type ref to cl_abap_zip,
          lo_worksheet        type ref to zcl_excel_worksheet,
          lo_active_worksheet type ref to zcl_excel_worksheet,
          lo_iterator         type ref to zcl_excel_collection_iterator,
          lo_nested_iterator  type ref to zcl_excel_collection_iterator,
          lo_table            type ref to zcl_excel_table,
          lo_drawing          type ref to zcl_excel_drawing,
          lo_drawings         type ref to zcl_excel_drawings.

    data: lv_content         type xstring,
          lv_active          type abap_boolean,
          lv_xl_sheet        type string,
          lv_xl_sheet_rels   type string,
          lv_xl_drawing      type string,
          lv_xl_drawing_rels type string,
          lv_syindex         type string,
          lv_value           type string,
          lv_drawing_index   type i,
          lv_comment_index   type i. " (+) Issue 588

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
    lv_drawing_index = 1.

    while lo_iterator->has_next( ) eq abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      if lo_active_worksheet->get_guid( ) eq lo_worksheet->get_guid( ).
        lv_active = abap_true.
      else.
        lv_active = abap_false.
      endif.

      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.
      lv_syindex = sy-index.
      lv_comment_index = sy-index. " (+) Issue 588
      shift lv_syindex right deleting trailing space.
      shift lv_syindex left deleting leading space.
      replace all occurrences of '#' in lv_xl_sheet with lv_syindex.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index ). " (+) Issue 588
      replace all occurrences of '#' in lv_xl_sheet_rels with lv_syindex.
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

* Add drawings **********************************
      lo_drawings = lo_worksheet->get_drawings( ).
      if lo_drawings->is_empty( ) = abap_false.
        lv_syindex = lv_drawing_index.
        shift lv_syindex right deleting trailing space.
        shift lv_syindex left deleting leading space.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        replace all occurrences of '#' in lv_xl_drawing with lv_syindex.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        replace all occurrences of '#' in lv_xl_drawing_rels with lv_syindex.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
        lv_drawing_index += 1.
      endif.
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
      lv_value = lo_drawing->get_media_name( ).
      concatenate 'xl/charts/' lv_value into lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    endwhile.

**********************************************************************
* STEP 9: Add vbaProject.bin to zip
    lo_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

**********************************************************************
* STEP 12: Create the final zip
    ep_excel = lo_zip->save( ).

  endmethod.


  method create_content_types.
** Constant node name
    data: lc_xml_node_workb_ct    type string value 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
          lc_xml_node_default     type string value 'Default',
          " Node attributes
          lc_xml_attr_partname    type string value 'PartName',
          lc_xml_attr_extension   type string value 'Extension',
          lc_xml_attr_contenttype type string value 'ContentType',
          lc_xml_node_workb_pn    type string value '/xl/workbook.xml',
          lc_xml_node_bin_ext     type string value 'bin',
          lc_xml_node_bin_ct      type string value 'application/vnd.ms-office.vbaProject'.


    data: lo_ixml          type ref to if_ixml_core,
          lo_document      type ref to if_ixml_document,
          lo_parser        type ref to if_ixml_parser_core,
          lo_element_root  type ref to if_ixml_node,
          lo_element       type ref to if_ixml_element,
          lo_collection    type ref to if_ixml_node_collection,
          lo_iterator      type ref to if_ixml_node_iterator,
          lo_streamfactory type ref to if_ixml_stream_factory_core,
          lo_istream       type ref to if_ixml_istream_core,
          lo_ostream       type ref to if_ixml_ostream_core,
          lo_renderer      type ref to if_ixml_renderer_core.

    data: lv_contenttype type string.

**********************************************************************
* STEP 3: Create standard contentType
    ep_content = super->create_content_types( ).

**********************************************************************
* STEP 2: modify XML adding the extension bin definition

    lo_ixml = cl_ixml_core=>create( ).
    lo_document      = lo_ixml->create_document( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_istream = lo_streamfactory->create_istream_xstring( string = ep_content ).
    lo_parser = lo_ixml->create_parser( istream = lo_istream document = lo_document stream_factory = lo_streamfactory ).
    lo_parser->parse( ).

    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_bin_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_bin_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: modify XML changing the contentType of node Override /xl/workbook.xml

    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    while lo_element is bound.
      lv_contenttype = lo_element->get_attribute_ns( lc_xml_attr_partname ).
      if lv_contenttype eq lc_xml_node_workb_pn.
        lo_element->remove_attribute_ns( lc_xml_attr_contenttype ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_workb_ct ).
        exit.
      endif.
      lo_element ?= lo_iterator->get_next( ).
    endwhile.

**********************************************************************
* STEP 3: Create xstring stream
    clear ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  endmethod.


  method create_xl_relationships.

** Constant node name
    data: lc_xml_node_relationship type string value 'Relationship',
          " Node attributes
          lc_xml_attr_id           type string value 'Id',
          lc_xml_attr_type         type string value 'Type',
          lc_xml_attr_target       type string value 'Target',
          " Node id
          lc_xml_node_ridx_id      type string value 'rId#',
          " Node type
          lc_xml_node_rid_vba_tp   type string value 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
          " Node target
          lc_xml_node_rid_vba_tg   type string value 'vbaProject.bin'.

    data: lo_ixml          type ref to if_ixml_core,
          lo_document      type ref to if_ixml_document,
          lo_parser        type ref to if_ixml_parser_core,
          lo_element_root  type ref to if_ixml_node,
          lo_element       type ref to if_ixml_element,
          lo_streamfactory type ref to if_ixml_stream_factory_core,
          lo_ostream       type ref to if_ixml_ostream_core,
          lo_istream       type ref to if_ixml_istream_core,
          lo_renderer      type ref to if_ixml_renderer_core.

    data: lv_xml_node_ridx_id type string,
          lv_size             type i,
          lv_syindex(2)       type c.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_relationships( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation


    lo_ixml = cl_ixml_core=>create( ).
    lo_document      = lo_ixml->create_document( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_istream = lo_streamfactory->create_istream_xstring( string = ep_content ).
    lo_parser = lo_ixml->create_parser( istream = lo_istream document = lo_document stream_factory = lo_streamfactory ).
    lo_parser->parse( ).

    lo_element_root = lo_document->if_ixml_node~get_first_child( ).


    lv_size = excel->get_worksheets_size( ).

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size += 4.
    lv_syindex = lv_size.
    shift lv_syindex right deleting trailing space.
    shift lv_syindex left deleting leading space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    replace all occurrences of '#' in lv_xml_node_ridx_id with lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_vba_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_vba_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: Create xstring stream
    clear ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  endmethod.


  method create_xl_sheet.

** Constant node name
    data: lc_xml_attr_codename      type string value 'codeName'.

    data: lo_ixml          type ref to if_ixml_core,
          lo_document      type ref to if_ixml_document,
          lo_element_root  type ref to if_ixml_node,
          lo_element       type ref to if_ixml_element,
          lo_collection    type ref to if_ixml_node_collection,
          lo_iterator      type ref to if_ixml_node_iterator,
          lo_streamfactory type ref to if_ixml_stream_factory_core,
          lo_istream       type ref to if_ixml_istream_core,
          lo_parser        type ref to if_ixml_parser_core,
          lo_ostream       type ref to if_ixml_ostream_core,
          lo_renderer      type ref to if_ixml_renderer_core.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_sheet( io_worksheet = io_worksheet
                                         iv_active    = iv_active ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    lo_ixml = cl_ixml_core=>create( ).
    lo_document      = lo_ixml->create_document( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_istream = lo_streamfactory->create_istream_xstring( string = ep_content ).
    lo_parser = lo_ixml->create_parser( istream = lo_istream document = lo_document stream_factory = lo_streamfactory ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'sheetPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    while lo_element is bound.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = io_worksheet->zif_excel_sheet_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    endwhile.

**********************************************************************
* STEP 3: Create xstring stream
    clear ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  endmethod.


  method create_xl_workbook.

** Constant node name
    data: lc_xml_attr_codename      type string value 'codeName'.

    data: lo_ixml          type ref to if_ixml_core,
          lo_document      type ref to if_ixml_document,
          lo_istream       type ref to if_ixml_istream_core,
          lo_parser        type ref to if_ixml_parser_core,
          lo_element_root  type ref to if_ixml_node,
          lo_element       type ref to if_ixml_element,
          lo_collection    type ref to if_ixml_node_collection,
          lo_iterator      type ref to if_ixml_node_iterator,
          lo_streamfactory type ref to if_ixml_stream_factory_core,
          lo_ostream       type ref to if_ixml_ostream_core,
          lo_renderer      type ref to if_ixml_renderer_core.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_workbook( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    lo_ixml = cl_ixml_core=>create( ).
    lo_document      = lo_ixml->create_document( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_istream = lo_streamfactory->create_istream_xstring( string = ep_content ).
    lo_parser = lo_ixml->create_parser( istream = lo_istream document = lo_document stream_factory = lo_streamfactory ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'fileVersion' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    while lo_element is bound.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename ).
      lo_element ?= lo_iterator->get_next( ).
    endwhile.

    lo_collection = lo_document->get_elements_by_tag_name( 'workbookPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    while lo_element is bound.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    endwhile.

**********************************************************************
* STEP 3: Create xstring stream
    clear ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  endmethod.
endclass.
