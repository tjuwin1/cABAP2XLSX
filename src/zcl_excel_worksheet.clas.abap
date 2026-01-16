class zcl_excel_worksheet definition
  public
  create public .

  public section.
*"* public components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!

    interfaces zif_excel_sheet_printsettings .
    interfaces zif_excel_sheet_properties .
    interfaces zif_excel_sheet_protection .
    interfaces zif_excel_sheet_vba_project .

    types:
      begin of  mty_s_outline_row,
        row_from  type i,
        row_to    type i,
        collapsed type abap_bool,
      end of mty_s_outline_row .
    types: mty_ts_outlines_row type sorted table of mty_s_outline_row with unique key primary_key components row_from row_to
                                                                      with non-unique sorted key row_to components row_to collapsed.
    types:
      begin of mty_s_ignored_errors,
        "! Cell reference (e.g. "A1") or list like "A1 A2" or range "A1:G1"
        cell_coords           type zif_excel_data_decl=>zexcel_cell_coords,
        "! Ignore errors when cells contain formulas that result in an error.
        eval_error            type abap_bool,
        "! Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        two_digit_text_year   type abap_bool,
        "! Ignore errors when numbers are formatted as text or are preceded by an apostrophe.
        number_stored_as_text type abap_bool,
        "! Ignore errors when a formula in a region of your WorkSheet differs from other formulas in the same region.
        formula               type abap_bool,
        "! Ignore errors when formulas omit certain cells in a region.
        formula_range         type abap_bool,
        "! Ignore errors when unlocked cells contain formulas.
        unlocked_formula      type abap_bool,
        "! Ignore errors when formulas refer to empty cells.
        empty_cell_reference  type abap_bool,
        "! Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified.
        list_data_validation  type abap_bool,
        "! Ignore errors when cells contain a value different from a calculated column formula.
        "! In other words, for a calculated column, a cell in that column is considered to have an error
        "! if its formula is different from the calculated column formula, or doesn't contain a formula at all.
        calculated_column     type abap_bool,
      end of mty_s_ignored_errors .
    types:
      mty_th_ignored_errors type hashed table of mty_s_ignored_errors with unique key cell_coords .
    types:
      begin of mty_s_column_formula,
        id                     type i,
        column                 type zif_excel_data_decl=>zexcel_cell_column,
        formula                type string,
        table_top_left_row     type zif_excel_data_decl=>zexcel_cell_row,
        table_bottom_right_row type zif_excel_data_decl=>zexcel_cell_row,
        table_left_column_int  type zif_excel_data_decl=>zexcel_cell_column,
        table_right_column_int type zif_excel_data_decl=>zexcel_cell_column,
      end of mty_s_column_formula .
    types:
      mty_th_column_formula
               type hashed table of mty_s_column_formula
               with unique key id .
    types:
      ty_doc_url type c length 255 .
    types:
      begin of mty_merge,
        row_from type i,
        row_to   type i,
        col_from type i,
        col_to   type i,
      end of mty_merge .
    types:
        mty_ts_merge type sorted table of mty_merge with unique key table_line.

    types:
      ty_area type c length 1 .

    constants c_break_column type zif_excel_data_decl=>zexcel_break value 2 ##NO_TEXT.
    constants c_break_none type zif_excel_data_decl=>zexcel_break value 0 ##NO_TEXT.
    constants c_break_row type zif_excel_data_decl=>zexcel_break value 1 ##NO_TEXT.
    constants:
      begin of c_area,
        whole   type ty_area value 'W',                     "#EC NOTEXT
        topleft type ty_area value 'T',                     "#EC NOTEXT
      end of c_area .
    data excel type ref to zcl_excel read-only .
    data print_gridlines type zif_excel_data_decl=>zexcel_print_gridlines read-only value abap_false ##NO_TEXT.
    data sheet_content type zif_excel_data_decl=>zexcel_t_cell_data .
    data sheet_setup type ref to zcl_excel_sheet_setup .
    data show_gridlines type zif_excel_data_decl=>zexcel_show_gridlines read-only value abap_true ##NO_TEXT.
    data show_rowcolheaders type zif_excel_data_decl=>zexcel_show_gridlines read-only value abap_true ##NO_TEXT.
    data tabcolor type zif_excel_data_decl=>zexcel_s_tabcolor read-only .
    data column_formulas type mty_th_column_formula read-only .
    class-data:
      begin of c_messages read-only,
        formula_id_only_is_possible type string,
        column_formula_id_not_found type string,
        formula_not_in_this_table   type string,
        formula_in_other_column     type string,
      end of c_messages .
    data mt_merged_cells type mty_ts_merge read-only .
    data pane_top_left_cell type string read-only.
    data sheetview_top_left_cell type string read-only.

    methods add_comment
      importing
        !ip_comment type ref to zcl_excel_comment .
    methods add_drawing
      importing
        !ip_drawing type ref to zcl_excel_drawing .
    methods add_new_column
      importing
        !ip_column       type simple
      returning
        value(eo_column) type ref to zcl_excel_column
      raising
        zcx_excel .
    methods add_new_style_cond
      importing
        !ip_dimension_range  type string default 'A1'
      returning
        value(eo_style_cond) type ref to zcl_excel_style_cond .
    methods add_new_data_validation
      returning
        value(eo_data_validation) type ref to zcl_excel_data_validation .
    methods add_new_range
      returning
        value(eo_range) type ref to zcl_excel_range .
    methods add_new_row
      importing
        !ip_row       type simple
      returning
        value(eo_row) type ref to zcl_excel_row .
*    METHODS bind_alv
*      IMPORTING
*        !io_alv      TYPE REF TO object
*        !it_table    TYPE STANDARD TABLE
*        !i_top       TYPE i DEFAULT 1
*        !i_left      TYPE i DEFAULT 1
*        !table_style type zif_excel_data_decl=>zexcel_table_style OPTIONAL
*        !i_table     TYPE abap_bool DEFAULT abap_true
*      RAISING
*        zcx_excel .
*    METHODS bind_alv_ole2
*      IMPORTING
*        !i_document_url      TYPE ty_doc_url DEFAULT space
*        !i_xls               TYPE c DEFAULT space
*        !i_save_path         TYPE string
*        !io_alv              TYPE REF TO cl_gui_alv_grid
*        !it_listheader       TYPE slis_t_listheader OPTIONAL
*        !i_top               TYPE i DEFAULT 1
*        !i_left              TYPE i DEFAULT 1
*        !i_columns_header    TYPE c DEFAULT 'X'
*        !i_columns_autofit   TYPE c DEFAULT 'X'
*        !i_format_col_header TYPE soi_format_item OPTIONAL
*        !i_format_subtotal   TYPE soi_format_item OPTIONAL
*        !i_format_total      TYPE soi_format_item OPTIONAL
*      EXCEPTIONS
*        miss_guide
*        ex_transfer_kkblo_error
*        fatal_error
*        inv_data_range
*        dim_mismatch_vkey
*        dim_mismatch_sema
*        error_in_sema .
    methods bind_table
      importing
        !ip_table               type standard table
        !it_field_catalog       type zif_excel_data_decl=>zexcel_t_fieldcatalog optional
        !is_table_settings      type zif_excel_data_decl=>zexcel_s_table_settings optional
        value(iv_default_descr) type c optional
        !iv_no_line_if_empty    type abap_bool default abap_false
      exporting
        !es_table_settings      type zif_excel_data_decl=>zexcel_s_table_settings
      raising
        zcx_excel .
    methods calculate_column_widths
      raising
        zcx_excel .
    methods change_area_style
      importing
        !ip_range         type csequence optional
        !ip_column_start  type simple optional
        !ip_column_end    type simple optional
        !ip_row           type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to        type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_style_changer type ref to zif_excel_style_changer
      raising
        zcx_excel .
    methods change_cell_style
      importing
        !ip_columnrow                   type csequence optional
        !ip_column                      type simple optional
        !ip_row                         type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_complete                    type zif_excel_data_decl=>zexcel_s_cstyle_complete optional
        !ip_xcomplete                   type zif_excel_data_decl=>zexcel_s_cstylex_complete optional
        !ip_font                        type zif_excel_data_decl=>zexcel_s_cstyle_font optional
        !ip_xfont                       type zif_excel_data_decl=>zexcel_s_cstylex_font optional
        !ip_fill                        type zif_excel_data_decl=>zexcel_s_cstyle_fill optional
        !ip_xfill                       type zif_excel_data_decl=>zexcel_s_cstylex_fill optional
        !ip_borders                     type zif_excel_data_decl=>zexcel_s_cstyle_borders optional
        !ip_xborders                    type zif_excel_data_decl=>zexcel_s_cstylex_borders optional
        !ip_alignment                   type zif_excel_data_decl=>zexcel_s_cstyle_alignment optional
        !ip_xalignment                  type zif_excel_data_decl=>zexcel_s_cstylex_alignment optional
        !ip_number_format_format_code   type zif_excel_data_decl=>zexcel_number_format optional
        !ip_protection                  type zif_excel_data_decl=>zexcel_s_cstyle_protection optional
        !ip_xprotection                 type zif_excel_data_decl=>zexcel_s_cstylex_protection optional
        !ip_font_bold                   type abap_boolean optional
        !ip_font_color                  type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_font_color_rgb              type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_font_color_indexed          type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_font_color_theme            type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_font_color_tint             type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_font_family                 type zif_excel_data_decl=>zexcel_style_font_family optional
        !ip_font_italic                 type abap_boolean optional
        !ip_font_name                   type zif_excel_data_decl=>zexcel_style_font_name optional
        !ip_font_scheme                 type zif_excel_data_decl=>zexcel_style_font_scheme optional
        !ip_font_size                   type zif_excel_data_decl=>zexcel_style_font_size optional
        !ip_font_strikethrough          type abap_boolean optional
        !ip_font_underline              type abap_boolean optional
        !ip_font_underline_mode         type zif_excel_data_decl=>zexcel_style_font_underline optional
        !ip_fill_filltype               type zif_excel_data_decl=>zexcel_fill_type optional
        !ip_fill_rotation               type zif_excel_data_decl=>zexcel_rotation optional
        !ip_fill_fgcolor                type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_fill_fgcolor_rgb            type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_fill_fgcolor_indexed        type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_fill_fgcolor_theme          type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_fill_fgcolor_tint           type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_fill_bgcolor                type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_fill_bgcolor_rgb            type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_fill_bgcolor_indexed        type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_fill_bgcolor_theme          type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_fill_bgcolor_tint           type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_allborders          type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_fill_gradtype_type          type zif_excel_data_decl=>zexcel_s_gradient_type-type optional
        !ip_fill_gradtype_degree        type zif_excel_data_decl=>zexcel_s_gradient_type-degree optional
        !ip_xborders_allborders         type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_borders_diagonal            type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_fill_gradtype_bottom        type zif_excel_data_decl=>zexcel_s_gradient_type-bottom optional
        !ip_fill_gradtype_top           type zif_excel_data_decl=>zexcel_s_gradient_type-top optional
        !ip_xborders_diagonal           type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_borders_diagonal_mode       type zif_excel_data_decl=>zexcel_diagonal optional
        !ip_fill_gradtype_right         type zif_excel_data_decl=>zexcel_s_gradient_type-right optional
        !ip_borders_down                type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_fill_gradtype_left          type zif_excel_data_decl=>zexcel_s_gradient_type-left optional
        !ip_fill_gradtype_position1     type zif_excel_data_decl=>zexcel_s_gradient_type-position1 optional
        !ip_xborders_down               type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_borders_left                type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_fill_gradtype_position2     type zif_excel_data_decl=>zexcel_s_gradient_type-position2 optional
        !ip_fill_gradtype_position3     type zif_excel_data_decl=>zexcel_s_gradient_type-position3 optional
        !ip_xborders_left               type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_borders_right               type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_xborders_right              type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_borders_top                 type zif_excel_data_decl=>zexcel_s_cstyle_border optional
        !ip_xborders_top                type zif_excel_data_decl=>zexcel_s_cstylex_border optional
        !ip_alignment_horizontal        type zif_excel_data_decl=>zexcel_alignment optional
        !ip_alignment_vertical          type zif_excel_data_decl=>zexcel_alignment optional
        !ip_alignment_textrotation      type zif_excel_data_decl=>zexcel_text_rotation optional
        !ip_alignment_wraptext          type abap_boolean optional
        !ip_alignment_shrinktofit       type abap_boolean optional
        !ip_alignment_indent            type zif_excel_data_decl=>zexcel_indent optional
        !ip_protection_hidden           type zif_excel_data_decl=>zexcel_cell_protection optional
        !ip_protection_locked           type zif_excel_data_decl=>zexcel_cell_protection optional
        !ip_borders_allborders_style    type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_allborders_color    type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_allbo_color_rgb     type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_allbo_color_indexed type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_allbo_color_theme   type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_allbo_color_tint    type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_diagonal_style      type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_diagonal_color      type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_diagonal_color_rgb  type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_diagonal_color_inde type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_diagonal_color_them type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_diagonal_color_tint type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_down_style          type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_down_color          type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_down_color_rgb      type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_down_color_indexed  type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_down_color_theme    type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_down_color_tint     type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_left_style          type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_left_color          type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_left_color_rgb      type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_left_color_indexed  type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_left_color_theme    type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_left_color_tint     type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_right_style         type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_right_color         type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_right_color_rgb     type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_right_color_indexed type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_right_color_theme   type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_right_color_tint    type zif_excel_data_decl=>zexcel_style_color_tint optional
        !ip_borders_top_style           type zif_excel_data_decl=>zexcel_border optional
        !ip_borders_top_color           type zif_excel_data_decl=>zexcel_s_style_color optional
        !ip_borders_top_color_rgb       type zif_excel_data_decl=>zexcel_style_color_argb optional
        !ip_borders_top_color_indexed   type zif_excel_data_decl=>zexcel_style_color_indexed optional
        !ip_borders_top_color_theme     type zif_excel_data_decl=>zexcel_style_color_theme optional
        !ip_borders_top_color_tint      type zif_excel_data_decl=>zexcel_style_color_tint optional
      returning
        value(ep_guid)                  type zif_excel_data_decl=>zexcel_cell_style
      raising
        zcx_excel .
    class-methods class_constructor .
    methods constructor
      importing
        !ip_excel type ref to zcl_excel
        !ip_title type zif_excel_data_decl=>zexcel_sheet_title optional
      raising
        zcx_excel .
    methods delete_merge
      importing
        !ip_cell_column type simple optional
        !ip_cell_row    type zif_excel_data_decl=>zexcel_cell_row optional
      raising
        zcx_excel .
    methods delete_row_outline
      importing
        !iv_row_from type i
        !iv_row_to   type i
      raising
        zcx_excel .
    methods freeze_panes
      importing
        !ip_num_columns type i optional
        !ip_num_rows    type i optional
      raising
        zcx_excel .
    methods get_active_cell
      returning
        value(ep_active_cell) type string
      raising
        zcx_excel .
    methods get_cell
      importing
        !ip_columnrow type csequence optional
        !ip_column    type simple optional
        !ip_row       type zif_excel_data_decl=>zexcel_cell_row optional
      exporting
        !ep_value     type zif_excel_data_decl=>zexcel_cell_value
        !ep_rc        type sysubrc
        !ep_style     type ref to zcl_excel_style
        !ep_guid      type zif_excel_data_decl=>zexcel_cell_style
        !ep_formula   type zif_excel_data_decl=>zexcel_cell_formula
        !et_rtf       type zif_excel_data_decl=>zexcel_t_rtf
      raising
        zcx_excel .
    methods get_column
      importing
        !ip_column       type simple
      returning
        value(eo_column) type ref to zcl_excel_column
      raising
        zcx_excel .
    methods get_columns
      returning
        value(eo_columns) type ref to zcl_excel_columns
      raising
        zcx_excel .
    methods get_columns_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator
      raising
        zcx_excel .
    methods get_style_cond_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_data_validations_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_data_validations_size
      returning
        value(ep_size) type i .
    methods get_default_column
      returning
        value(eo_column) type ref to zcl_excel_column
      raising
        zcx_excel .
    methods get_default_excel_date_format
      returning
        value(ep_default_excel_date_format) type zif_excel_data_decl=>zexcel_number_format .
    methods get_default_excel_time_format
      returning
        value(ep_default_excel_time_format) type zif_excel_data_decl=>zexcel_number_format .
    methods get_default_row
      returning
        value(eo_row) type ref to zcl_excel_row .
    methods get_dimension_range
      returning
        value(ep_dimension_range) type string
      raising
        zcx_excel .
    methods get_comments
      importing
        iv_copy_collection type abap_boolean default abap_true
      returning
        value(r_comments)  type ref to zcl_excel_comments .
    methods get_drawings
      importing
        !ip_type          type zif_excel_data_decl=>zexcel_drawing_type optional
      returning
        value(r_drawings) type ref to zcl_excel_drawings .
    methods get_comments_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_drawings_iterator
      importing
        !ip_type           type zif_excel_data_decl=>zexcel_drawing_type
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_freeze_cell
      exporting
        !ep_row    type zif_excel_data_decl=>zexcel_cell_row
        !ep_column type zif_excel_data_decl=>zexcel_cell_column .
    methods get_guid
      returning
        value(ep_guid) type sysuuid_x16 .
    methods get_highest_column
      returning
        value(r_highest_column) type zif_excel_data_decl=>zexcel_cell_column
      raising
        zcx_excel .
    methods get_highest_row
      returning
        value(r_highest_row) type int4
      raising
        zcx_excel .
    methods get_hyperlinks_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_hyperlinks_size
      returning
        value(ep_size) type i .
    methods get_ignored_errors
      returning
        value(rt_ignored_errors) type mty_th_ignored_errors .
    methods get_merge
      returning
        value(merge_range) type string_table
      raising
        zcx_excel .
    methods get_pagebreaks
      returning
        value(ro_pagebreaks) type ref to zcl_excel_worksheet_pagebreaks
      raising
        zcx_excel .
    methods get_ranges_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_row
      importing
        !ip_row       type int4
      returning
        value(eo_row) type ref to zcl_excel_row .
    methods get_rows
      returning
        value(eo_rows) type ref to zcl_excel_rows .
    methods get_rows_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_row_outlines
      returning
        value(rt_row_outlines) type mty_ts_outlines_row .
    methods get_style_cond
      importing
        !ip_guid             type zif_excel_data_decl=>zexcel_cell_style
      returning
        value(eo_style_cond) type ref to zcl_excel_style_cond .
    methods get_tabcolor
      returning
        value(ev_tabcolor) type zif_excel_data_decl=>zexcel_s_tabcolor .
    methods get_tables_iterator
      returning
        value(eo_iterator) type ref to zcl_excel_collection_iterator .
    methods get_tables_size
      returning
        value(ep_size) type i .
    methods get_title
      importing
        !ip_escaped     type abap_boolean default ''
      returning
        value(ep_title) type zif_excel_data_decl=>zexcel_sheet_title .
    methods is_cell_merged
      importing
        !ip_column          type simple
        !ip_row             type zif_excel_data_decl=>zexcel_cell_row
      returning
        value(rp_is_merged) type abap_bool
      raising
        zcx_excel .
    methods set_cell
      importing
        !ip_columnrow         type csequence optional
        !ip_column            type simple optional
        !ip_row               type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_value             type simple optional
        !ip_formula           type zif_excel_data_decl=>zexcel_cell_formula optional
        !ip_style             type any optional
        !ip_hyperlink         type ref to zcl_excel_hyperlink optional
        !ip_data_type         type zif_excel_data_decl=>zexcel_cell_data_type optional
        !ip_abap_type         type abap_typekind optional
        !ip_currency          type waers_curc optional
        !ip_textvalue         type csequence optional
        !ip_unitofmeasure     type meins optional
        !it_rtf               type zif_excel_data_decl=>zexcel_t_rtf optional
        !ip_column_formula_id type mty_s_column_formula-id optional
      raising
        zcx_excel .
    methods set_cell_formula
      importing
        !ip_columnrow type csequence optional
        !ip_column    type simple optional
        !ip_row       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_formula   type zif_excel_data_decl=>zexcel_cell_formula
      raising
        zcx_excel .
    methods set_cell_style
      importing
        !ip_columnrow type csequence optional
        !ip_column    type simple optional
        !ip_row       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_style     type any
      raising
        zcx_excel .
    methods set_column_width
      importing
        !ip_column         type simple
        !ip_width_fix      type simple default 0
        !ip_width_autosize type abap_boolean default 'X'
      raising
        zcx_excel .
    methods set_default_excel_date_format
      importing
        !ip_default_excel_date_format type zif_excel_data_decl=>zexcel_number_format
      raising
        zcx_excel .
    methods set_ignored_errors
      importing
        !it_ignored_errors type mty_th_ignored_errors .
    methods set_merge
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_style        type any optional
        !ip_value        type simple optional          "added parameter
        !ip_formula      type zif_excel_data_decl=>zexcel_cell_formula optional        "added parameter
      raising
        zcx_excel .
    methods set_pane_top_left_cell
      importing
        !iv_columnrow type csequence
      raising
        zcx_excel.
    methods set_print_gridlines
      importing
        !i_print_gridlines type zif_excel_data_decl=>zexcel_print_gridlines .
    methods set_row_height
      importing
        !ip_row        type simple
        !ip_height_fix type simple
      raising
        zcx_excel .
    methods set_row_outline
      importing
        !iv_row_from  type i
        !iv_row_to    type i
        !iv_collapsed type abap_bool
      raising
        zcx_excel .
    methods set_sheetview_top_left_cell
      importing
        !iv_columnrow type csequence
      raising
        zcx_excel.
    methods set_show_gridlines
      importing
        !i_show_gridlines type zif_excel_data_decl=>zexcel_show_gridlines .
    methods set_show_rowcolheaders
      importing
        !i_show_rowcolheaders type zif_excel_data_decl=>zexcel_show_rowcolheader .
    methods set_tabcolor
      importing
        !iv_tabcolor type zif_excel_data_decl=>zexcel_s_tabcolor .
    methods set_table
      importing
        !ip_table           type standard table
        !ip_hdr_style       type any optional
        !ip_body_style      type any optional
        !ip_table_title     type string
        !ip_top_left_column type zif_excel_data_decl=>zexcel_cell_column_alpha default 'B'
        !ip_top_left_row    type zif_excel_data_decl=>zexcel_cell_row default 3
        !ip_transpose       type abap_bool optional
        !ip_no_header       type abap_bool optional
      raising
        zcx_excel .
    methods set_title
      importing
        !ip_title type zif_excel_data_decl=>zexcel_sheet_title
      raising
        zcx_excel .
    methods get_table
      importing
        !iv_skipped_rows           type int4 default 0
        !iv_skipped_cols           type int4 default 0
        !iv_max_col                type int4 optional
        !iv_max_row                type int4 optional
        !iv_skip_bottom_empty_rows type abap_bool default abap_false
      exporting
        !et_table                  type standard table
      raising
        zcx_excel .
    methods set_merge_style
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_style        type any optional
      raising
        zcx_excel .
    methods set_area_formula
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_formula      type zif_excel_data_decl=>zexcel_cell_formula
        !ip_merge        type abap_bool optional
        !ip_area         type ty_area default c_area-topleft
      raising
        zcx_excel .
    methods set_area_style
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_style        type any
        !ip_merge        type abap_bool optional
      raising
        zcx_excel .
    methods set_area
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_value        type simple optional
        !ip_formula      type zif_excel_data_decl=>zexcel_cell_formula optional
        !ip_style        type any optional
        !ip_hyperlink    type ref to zcl_excel_hyperlink optional
        !ip_data_type    type zif_excel_data_decl=>zexcel_cell_data_type optional
        !ip_abap_type    type abap_typekind optional
        !ip_merge        type abap_bool optional
        !ip_area         type ty_area default c_area-topleft
      raising
        zcx_excel .
    methods get_header_footer_drawings
      returning
        value(rt_drawings) type zif_excel_data_decl=>zexcel_t_drawings .
    methods set_area_hyperlink
      importing
        !ip_range        type csequence optional
        !ip_column_start type simple optional
        !ip_column_end   type simple optional
        !ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
        !ip_url          type string
        !ip_is_internal  type abap_bool
      raising
        zcx_excel .
    "! excel upload, counterpart to BIND_TABLE
    "! @parameter it_field_catalog | field catalog, used to derive correct types
    "! @parameter iv_begin_row | starting row, by default 2 to skip header
    "! @parameter et_data | generic internal table, there may be conversion losses
    "! @parameter er_data | ref to internal table of string columns, to get raw data without conversion losses.
    methods convert_to_table
      importing
        !it_field_catalog type zif_excel_data_decl=>zexcel_t_fieldcatalog optional
        !iv_begin_row     type int4 default 2
        !iv_end_row       type int4 default 0
      exporting
        !et_data          type standard table
        !er_data          type ref to data
      raising
        zcx_excel .
  protected section.
    methods set_table_reference
      importing
        !ip_column    type zif_excel_data_decl=>zexcel_cell_column
        !ip_row       type zif_excel_data_decl=>zexcel_cell_row
        !ir_table     type ref to zcl_excel_table
        !ip_fieldname type zif_excel_data_decl=>zexcel_fieldname
        !ip_header    type abap_bool
      raising
        zcx_excel .
  private section.

*"* private components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
    types ty_table_settings type standard table of zif_excel_data_decl=>zexcel_s_table_settings with default key.

    constants typekind_utclong type abap_typekind value 'p'.

    class-data variable_utclong type ref to data.

    data active_cell type zif_excel_data_decl=>zexcel_s_cell_data .
    data charts type ref to zcl_excel_drawings .
    data columns type ref to zcl_excel_columns .
    data row_default type ref to zcl_excel_row .
    data column_default type ref to zcl_excel_column .
    data styles_cond type ref to zcl_excel_styles_cond .
    data data_validations type ref to zcl_excel_data_validations .
    data default_excel_date_format type zif_excel_data_decl=>zexcel_number_format .
    data default_excel_time_format type zif_excel_data_decl=>zexcel_number_format .
    data comments type ref to zcl_excel_comments .
    data drawings type ref to zcl_excel_drawings .
    data freeze_pane_cell_column type zif_excel_data_decl=>zexcel_cell_column .
    data freeze_pane_cell_row type zif_excel_data_decl=>zexcel_cell_row .
    data guid type sysuuid_x16 .
    data hyperlinks type ref to zcl_excel_collection .
    data lower_cell type zif_excel_data_decl=>zexcel_s_cell_data .
    data mo_pagebreaks type ref to zcl_excel_worksheet_pagebreaks .
    data mt_row_outlines type mty_ts_outlines_row .
    data print_title_col_from type zif_excel_data_decl=>zexcel_cell_column_alpha .
    data print_title_col_to type zif_excel_data_decl=>zexcel_cell_column_alpha .
    data print_title_row_from type zif_excel_data_decl=>zexcel_cell_row .
    data print_title_row_to type zif_excel_data_decl=>zexcel_cell_row .
    data ranges type ref to zcl_excel_ranges .
    data rows type ref to zcl_excel_rows .
    data tables type ref to zcl_excel_collection .
    data title type zif_excel_data_decl=>zexcel_sheet_title value 'Worksheet'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    data upper_cell type zif_excel_data_decl=>zexcel_s_cell_data .
    data mt_ignored_errors type mty_th_ignored_errors.
    data right_to_left type abap_bool.

    methods calculate_cell_width
      importing
        !ip_column      type simple
        !ip_row         type zif_excel_data_decl=>zexcel_cell_row
      returning
        value(ep_width) type f
      raising
        zcx_excel .
    class-methods calculate_table_bottom_right
      importing
        ip_table         type standard table
        it_field_catalog type zif_excel_data_decl=>zexcel_t_fieldcatalog
      changing
        cs_settings      type zif_excel_data_decl=>zexcel_s_table_settings
      raising
        zcx_excel.
    class-methods check_cell_column_formula
      importing
        it_column_formulas   type mty_th_column_formula
        ip_column_formula_id type mty_s_column_formula-id
        ip_formula           type zif_excel_data_decl=>zexcel_cell_formula
        ip_value             type simple
        ip_row               type zif_excel_data_decl=>zexcel_cell_row
        ip_column            type zif_excel_data_decl=>zexcel_cell_column
      raising
        zcx_excel.
    methods check_rtf
      importing
        !ip_value       type simple
        value(ip_style) type zif_excel_data_decl=>zexcel_cell_style optional
      changing
        !ct_rtf         type zif_excel_data_decl=>zexcel_t_rtf
      raising
        zcx_excel .
    class-methods check_table_overlapping
      importing
        is_table_settings       type zif_excel_data_decl=>zexcel_s_table_settings
        it_other_table_settings type ty_table_settings
      raising
        zcx_excel.
    methods clear_initial_colorxfields
      importing
        is_color  type zif_excel_data_decl=>zexcel_s_style_color
      changing
        cs_xcolor type zif_excel_data_decl=>zexcel_s_cstylex_color.
    methods generate_title
      returning
        value(ep_title) type zif_excel_data_decl=>zexcel_sheet_title .
    methods get_value_type
      importing
        !ip_value      type simple
      exporting
        !ep_value      type simple
        !ep_value_type type abap_typekind .
    methods move_supplied_borders
      importing
        iv_border_supplied        type abap_bool
        is_border                 type zif_excel_data_decl=>zexcel_s_cstyle_border
        iv_xborder_supplied       type abap_bool
        is_xborder                type zif_excel_data_decl=>zexcel_s_cstylex_border
      changing
        cs_complete_style_border  type zif_excel_data_decl=>zexcel_s_cstyle_border
        cs_complete_stylex_border type zif_excel_data_decl=>zexcel_s_cstylex_border.
    methods normalize_column_heading_texts
      importing
        iv_default_descr type c
        it_field_catalog type zif_excel_data_decl=>zexcel_t_fieldcatalog
      returning
        value(result)    type zif_excel_data_decl=>zexcel_t_fieldcatalog.
    methods normalize_columnrow_parameter
      importing
        ip_columnrow type csequence optional
        ip_column    type simple optional
        ip_row       type zif_excel_data_decl=>zexcel_cell_row optional
      exporting
        ep_column    type zif_excel_data_decl=>zexcel_cell_column
        ep_row       type zif_excel_data_decl=>zexcel_cell_row
      raising
        zcx_excel.
    methods normalize_range_parameter
      importing
        ip_range        type csequence optional
        ip_column_start type simple optional
        ip_column_end   type simple optional
        ip_row          type zif_excel_data_decl=>zexcel_cell_row optional
        ip_row_to       type zif_excel_data_decl=>zexcel_cell_row optional
      exporting
        ep_column_start type zif_excel_data_decl=>zexcel_cell_column
        ep_column_end   type zif_excel_data_decl=>zexcel_cell_column
        ep_row          type zif_excel_data_decl=>zexcel_cell_row
        ep_row_to       type zif_excel_data_decl=>zexcel_cell_row
      raising
        zcx_excel.
    class-methods normalize_style_parameter
      importing
        !ip_style_or_guid type any
      returning
        value(rv_guid)    type zif_excel_data_decl=>zexcel_cell_style
      raising
        zcx_excel .
    methods print_title_set_range .
    methods update_dimension_range
      raising
        zcx_excel .
endclass.



class zcl_excel_worksheet implementation.


  method add_comment.
    comments->include( ip_comment ).
  endmethod.                    "add_comment


  method add_drawing.
    case ip_drawing->get_type( ).
      when zcl_excel_drawing=>type_image.
        drawings->include( ip_drawing ).
      when zcl_excel_drawing=>type_chart.
        charts->include( ip_drawing ).
    endcase.
  endmethod.                    "ADD_DRAWING


  method add_new_column.
    data: lv_column_alpha type zif_excel_data_decl=>zexcel_cell_column_alpha.

    lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column ).

    create object eo_column
      exporting
        ip_index     = lv_column_alpha
        ip_excel     = me->excel
        ip_worksheet = me.
    columns->add( eo_column ).
  endmethod.                    "ADD_NEW_COLUMN


  method add_new_data_validation.

    create object eo_data_validation.
    data_validations->add( eo_data_validation ).
  endmethod.                    "ADD_NEW_DATA_VALIDATION


  method add_new_range.
* Create default blank range
    create object eo_range.
    ranges->add( eo_range ).
  endmethod.                    "ADD_NEW_RANGE


  method add_new_row.
    create object eo_row
      exporting
        ip_index = ip_row.
    rows->add( eo_row ).
  endmethod.                    "ADD_NEW_ROW


  method add_new_style_cond.
    create object eo_style_cond exporting ip_dimension_range = ip_dimension_range.
    styles_cond->add( eo_style_cond ).
*  ENDMETHOD.                    "ADD_NEW_STYLE_COND


*  METHOD bind_alv.
    "@TODO: Commented by Juwin
*    DATA: lo_converter TYPE REF TO zcl_excel_converter.
*
*    CREATE OBJECT lo_converter.
*
*    TRY.
*        lo_converter->convert(
*          EXPORTING
*            io_alv         = io_alv
*            it_table       = it_table
*            i_row_int      = i_top
*            i_column_int   = i_left
*            i_table        = i_table
*            i_style_table  = table_style
*            io_worksheet   = me
*          CHANGING
*            co_excel       = excel ).
*      CATCH zcx_excel .
*    ENDTRY.

*  ENDMETHOD.                    "BIND_ALV
*
*
*  METHOD bind_alv_ole2.
    "@TODO: Commented by Juwin
*    CALL METHOD ('ZCL_EXCEL_OLE')=>('BIND_ALV_OLE2')
*      EXPORTING
*        i_document_url          = i_document_url
*        i_xls                   = i_xls
*        i_save_path             = i_save_path
*        io_alv                  = io_alv
*        it_listheader           = it_listheader
*        i_top                   = i_top
*        i_left                  = i_left
*        i_columns_header        = i_columns_header
*        i_columns_autofit       = i_columns_autofit
*        i_format_col_header     = i_format_col_header
*        i_format_subtotal       = i_format_subtotal
*        i_format_total          = i_format_total
*      EXCEPTIONS
*        miss_guide              = 1
*        ex_transfer_kkblo_error = 2
*        fatal_error             = 3
*        inv_data_range          = 4
*        dim_mismatch_vkey       = 5
*        dim_mismatch_sema       = 6
*        error_in_sema           = 7
*        OTHERS                  = 8.
*    IF sy-subrc <> 0.
*      CASE sy-subrc.
*        WHEN 1. RAISE miss_guide.
*        WHEN 2. RAISE ex_transfer_kkblo_error.
*        WHEN 3. RAISE fatal_error.
*        WHEN 4. RAISE inv_data_range.
*        WHEN 5. RAISE dim_mismatch_vkey.
*        WHEN 6. RAISE dim_mismatch_sema.
*        WHEN 7. RAISE error_in_sema.
*      ENDCASE.
*    ENDIF.

  endmethod.                    "BIND_ALV_OLE2


  method bind_table.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmöcker,      (wi p)              2012-12-01
*              - ...
*          aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue #237   - Check if overlapping areas exist
*              - Alessandro Iannacci                        2012-12-01
* changes:     - Added raise if overlaps are detected
*--------------------------------------------------------------------*

    constants:
      lc_top_left_column type zif_excel_data_decl=>zexcel_cell_column_alpha value 'A',
      lc_top_left_row    type zif_excel_data_decl=>zexcel_cell_row value 1,
      lc_no_currency     type waers value is initial,
      lc_no_unit         type meins value is initial.

    data:
      lv_row_int              type zif_excel_data_decl=>zexcel_cell_row,
      lv_first_row            type zif_excel_data_decl=>zexcel_cell_row,
      lv_last_row             type zif_excel_data_decl=>zexcel_cell_row,
      lv_column_int           type zif_excel_data_decl=>zexcel_cell_column,
      lv_column_alpha         type zif_excel_data_decl=>zexcel_cell_column_alpha,
      lt_field_catalog        type zif_excel_data_decl=>zexcel_t_fieldcatalog,
      lv_id                   type i,
      lv_formula              type string,
      ls_settings             type zif_excel_data_decl=>zexcel_s_table_settings,
      lo_table                type ref to zcl_excel_table,
      lv_value_lowercase      type string,
      lv_syindex              type c length 3,
      lo_iterator             type ref to zcl_excel_collection_iterator,
      lo_style_cond           type ref to zcl_excel_style_cond,
      lo_curtable             type ref to zcl_excel_table,
      lt_other_table_settings type ty_table_settings.
    data: ls_column_formula type mty_s_column_formula,
          lv_mincol         type i.

    field-symbols:
      <ls_field_catalog>        type zif_excel_data_decl=>zexcel_s_fieldcatalog,
      <ls_field_catalog_custom> type zif_excel_data_decl=>zexcel_s_fieldcatalog,
      <fs_table_line>           type any,
      <fs_fldval>               type any,
      <fs_fldval_currency>      type waers,
      <fs_fldval_uom>           type meins,
      <fs_fldval_text>          type string.

    ls_settings = is_table_settings.

    if ls_settings-top_left_column is initial.
      ls_settings-top_left_column = lc_top_left_column.
    endif.

    if ls_settings-table_style is initial.
      ls_settings-table_style = zcl_excel_table=>builtinstyle_medium2.
    endif.

    if ls_settings-top_left_row is initial.
      ls_settings-top_left_row = lc_top_left_row.
    endif.

    if it_field_catalog is not supplied.
      lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = ip_table ).
    else.
      lt_field_catalog = it_field_catalog.
    endif.

    sort lt_field_catalog by position.

    calculate_table_bottom_right(
      exporting
        ip_table         = ip_table
        it_field_catalog = lt_field_catalog
      changing
        cs_settings      = ls_settings ).

* Check if overlapping areas exist

    lo_iterator = me->tables->get_iterator( ).
    while lo_iterator->has_next( ) eq abap_true.
      lo_curtable ?= lo_iterator->get_next( ).
      append lo_curtable->settings to lt_other_table_settings.
    endwhile.

    check_table_overlapping(
        is_table_settings       = ls_settings
        it_other_table_settings = lt_other_table_settings ).

* Start filling the table

    create object lo_table.
    lo_table->settings = ls_settings.
    lo_table->set_data( ir_data = ip_table ).
    lv_id = me->excel->get_next_table_id( ).
    lo_table->set_id( iv_id = lv_id ).

    me->tables->add( lo_table ).

    lv_column_int = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).
    lv_row_int = ls_settings-top_left_row.

    lt_field_catalog = normalize_column_heading_texts(
          iv_default_descr = iv_default_descr
          it_field_catalog = lt_field_catalog ).

* It is better to loop column by column (only visible column)
    loop at lt_field_catalog assigning <ls_field_catalog>.

      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      if <ls_field_catalog>-width is not initial.
        set_column_width( ip_column = lv_column_alpha ip_width_fix = <ls_field_catalog>-width ).
      endif.

      " First of all write column header
      if <ls_field_catalog>-style_header is not initial.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name
                      ip_style  = <ls_field_catalog>-style_header ).
      else.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name ).
      endif.

      me->set_table_reference( ip_column    = lv_column_int
                               ip_row       = lv_row_int
                               ir_table     = lo_table
                               ip_fieldname = <ls_field_catalog>-fieldname
                               ip_header    = abap_true ).

      if <ls_field_catalog>-column_formula is not initial.
        ls_column_formula-id                     = lines( column_formulas ) + 1.
        ls_column_formula-column                 = lv_column_int.
        ls_column_formula-formula                = <ls_field_catalog>-column_formula.
        ls_column_formula-table_top_left_row     = lo_table->settings-top_left_row.
        ls_column_formula-table_bottom_right_row = lo_table->settings-bottom_right_row.
        ls_column_formula-table_left_column_int  = lv_mincol.
        ls_column_formula-table_right_column_int = zcl_excel_common=>convert_column2int( lo_table->settings-bottom_right_column ).
        insert ls_column_formula into table column_formulas.
      endif.

      lv_row_int += 1.
      loop at ip_table assigning <fs_table_line>.

        assign component <ls_field_catalog>-fieldname of structure <fs_table_line> to <fs_fldval>.

        " issue #290 Add formula support in table
        if <ls_field_catalog>-formula eq abap_true.
          if <ls_field_catalog>-style is not initial.
            if <ls_field_catalog>-abap_type is not initial.
              me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_style    = <ls_field_catalog>-style ).
            else.
              me->set_cell( ip_column   = lv_column_alpha
                            ip_row      = lv_row_int
                            ip_formula  = <fs_fldval>
                            ip_style    = <ls_field_catalog>-style ).
            endif.
          elseif <ls_field_catalog>-abap_type is not initial.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type ).
          else.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval> ).
          endif.
        elseif <ls_field_catalog>-column_formula is not initial.
          " Column formulas
          if <ls_field_catalog>-style is not initial.
            if <ls_field_catalog>-abap_type is not initial.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_abap_type         = <ls_field_catalog>-abap_type
                            ip_style             = <ls_field_catalog>-style ).
            else.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_style             = <ls_field_catalog>-style ).
            endif.
          elseif <ls_field_catalog>-abap_type is not initial.
            me->set_cell( ip_column             = lv_column_alpha
                          ip_row                = lv_row_int
                          ip_column_formula_id  = ls_column_formula-id
                          ip_abap_type          = <ls_field_catalog>-abap_type ).
          else.
            me->set_cell( ip_column            = lv_column_alpha
                          ip_row               = lv_row_int
                          ip_column_formula_id = ls_column_formula-id ).
          endif.
        else.
          if <ls_field_catalog>-currency_column is initial.
            assign lc_no_currency to <fs_fldval_currency>.
          else.
            assign component <ls_field_catalog>-currency_column of structure <fs_table_line> to <fs_fldval_currency>.
          endif.
          if <ls_field_catalog>-unit_column is initial.
            assign lc_no_unit to <fs_fldval_uom>.
          else.
            assign component <ls_field_catalog>-unit_column of structure <fs_table_line> to <fs_fldval_uom>.
          endif.
          if <ls_field_catalog>-text_column is initial.
            assign `` to <fs_fldval_text>.
          else.
            assign component <ls_field_catalog>-text_column of structure <fs_table_line> to <fs_fldval_text>.
          endif.

          if <ls_field_catalog>-style is not initial.
            if <ls_field_catalog>-abap_type is not initial.
              me->set_cell( ip_column           = lv_column_alpha
                            ip_row              = lv_row_int
                            ip_value            = <fs_fldval>
                            ip_abap_type        = <ls_field_catalog>-abap_type
                            ip_currency         = <fs_fldval_currency>
                            ip_unitofmeasure = <fs_fldval_uom>
                            ip_textvalue     = <fs_fldval_text>
                            ip_style            = <ls_field_catalog>-style ).
            else.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_value  = <fs_fldval>
                            ip_currency = <fs_fldval_currency>
                            ip_unitofmeasure = <fs_fldval_uom>
                            ip_textvalue     = <fs_fldval_text>
                            ip_style  = <ls_field_catalog>-style ).
            endif.
          else.
            if <ls_field_catalog>-abap_type is not initial.
              me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_currency  = <fs_fldval_currency>
                            ip_unitofmeasure = <fs_fldval_uom>
                            ip_textvalue     = <fs_fldval_text>
                          ip_value  = <fs_fldval> ).
            else.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_currency = <fs_fldval_currency>
                            ip_unitofmeasure = <fs_fldval_uom>
                            ip_textvalue     = <fs_fldval_text>
                            ip_value  = <fs_fldval> ).
            endif.
          endif.
        endif.
        lv_row_int += 1.

      endloop.
      if sy-subrc <> 0 and iv_no_line_if_empty = abap_false. "create empty row if table has no data
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = space ).
        lv_row_int += 1.
      endif.

*--------------------------------------------------------------------*
      " totals
*--------------------------------------------------------------------*
      if <ls_field_catalog>-totals_function is not initial.
        lv_formula = lo_table->get_totals_formula( ip_column = <ls_field_catalog>-column_name ip_function = <ls_field_catalog>-totals_function ).
        if <ls_field_catalog>-style_total is not initial.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula
                        ip_style    = <ls_field_catalog>-style_total ).
        else.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula ).
        endif.
      endif.

      lv_row_int = ls_settings-top_left_row.
      lv_column_int += 1.

*--------------------------------------------------------------------*
      " conditional formatting
*--------------------------------------------------------------------*
      if <ls_field_catalog>-style_cond is not initial.
        lv_first_row    = ls_settings-top_left_row + 1. " +1 to exclude header
        lv_last_row     = ls_settings-bottom_right_row.
        lo_style_cond = me->get_style_cond( <ls_field_catalog>-style_cond ).
        lo_style_cond->set_range( ip_start_column  = lv_column_alpha
                                  ip_start_row     = lv_first_row
                                  ip_stop_column   = lv_column_alpha
                                  ip_stop_row      = lv_last_row ).
      endif.

    endloop.

*--------------------------------------------------------------------*
    " Set field catalog
*--------------------------------------------------------------------*
    lo_table->fieldcat = lt_field_catalog[].

    es_table_settings = ls_settings.
    es_table_settings-bottom_right_column = lv_column_alpha.
    " >> Issue #291
    if ip_table is initial.
      es_table_settings-bottom_right_row    = ls_settings-top_left_row + 2.           "Last rows
    else.
      es_table_settings-bottom_right_row    = ls_settings-bottom_right_row + 1. "Last rows
    endif.
    " << Issue #291

  endmethod.                    "BIND_TABLE


  method calculate_cell_width.
*--------------------------------------------------------------------*
* issue #293   - Roberto Bianco
*              - Christian Assig                            2014-03-14
*
* changes: - Calculate widths using SAPscript font metrics
*            (transaction SE73)
*          - Calculate the width of dates
*          - Add additional width for auto filter buttons
*          - Add cell padding to simulate Excel behavior
*--------------------------------------------------------------------*

    data: ld_cell_value                type zif_excel_data_decl=>zexcel_cell_value,
          ld_style_guid                type zif_excel_data_decl=>zexcel_cell_style,
          ls_stylemapping              type zif_excel_data_decl=>zexcel_s_stylemapping,
          lo_table_object              type ref to object,
          lo_table                     type ref to zcl_excel_table,
          ld_table_top_left_column     type zif_excel_data_decl=>zexcel_cell_column,
          ld_table_bottom_right_column type zif_excel_data_decl=>zexcel_cell_column,
          ld_flag_contains_auto_filter type abap_bool value abap_false,
          ld_flag_bold                 type abap_bool value abap_false,
          ld_flag_italic               type abap_bool value abap_false,
          ld_date                      type d,
          ld_date_char                 type c length 50,
          ld_time                      type t,
          ld_time_char                 type c length 20,
          ld_font_height               type zcl_excel_font=>ty_font_height value zcl_excel_font=>lc_default_font_height,
          ld_font_name                 type zif_excel_data_decl=>zexcel_style_font_name value zcl_excel_font=>lc_default_font_name.

    " Determine cell content and cell style
    me->get_cell( exporting ip_column = ip_column
                            ip_row    = ip_row
                  importing ep_value  = ld_cell_value
                            ep_guid   = ld_style_guid ).

    " ABAP2XLSX uses tables to define areas containing headers and
    " auto-filters. Find out if the current cell is in the header
    " of one of these tables.
    loop at me->tables->collection into lo_table_object.
      " Downcast: OBJECT -> ZCL_EXCEL_TABLE
      lo_table ?= lo_table_object.

      " Convert column letters to corresponding integer values
      ld_table_top_left_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-top_left_column ).

      ld_table_bottom_right_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-bottom_right_column ).

      " Is the current cell part of the table header?
      if ip_column between ld_table_top_left_column and
                           ld_table_bottom_right_column and
         ip_row    eq lo_table->settings-top_left_row.
        " Current cell is part of the table header
        " -> Assume that an auto filter is present and that the font is
        "    bold
        ld_flag_contains_auto_filter = abap_true.
        ld_flag_bold = abap_true.
      endif.
    endloop.

    " If a style GUID is present, read style attributes
    if ld_style_guid is not initial.
      try.
          " Read style attributes
          ls_stylemapping = me->excel->get_style_to_guid( ld_style_guid ).

          " If the current cell contains the default date format,
          " convert the cell value to a date and calculate its length
          case ls_stylemapping-complete_style-number_format-format_code.
            when zcl_excel_style_number_format=>c_format_date_std.

              " Convert excel date to ABAP date
              ld_date =
                zcl_excel_common=>excel_string_to_date( ld_cell_value ).

              " Format ABAP date using user's formatting settings
              ld_date_char = |{ ld_date date = user }|.

              " Remember the formatted date to calculate the cell size
              ld_cell_value = ld_date_char.

            when get_default_excel_time_format( ).

              ld_time = zcl_excel_common=>excel_string_to_time( ld_cell_value ).
              ld_time_char = |{ ld_time time = user }|.
              ld_cell_value = ld_time_char.

          endcase.


          " Read the font size and convert it to the font height
          " used by SAPscript (multiplication by 10)
          if ls_stylemapping-complete_stylex-font-size = abap_true.
            ld_font_height = ls_stylemapping-complete_style-font-size * 10.
          endif.

          " If set, remember the font name
          if ls_stylemapping-complete_stylex-font-name = abap_true.
            ld_font_name = ls_stylemapping-complete_style-font-name.
          endif.

          " If set, remember whether font is bold and italic.
          if ls_stylemapping-complete_stylex-font-bold = abap_true.
            ld_flag_bold = ls_stylemapping-complete_style-font-bold.
          endif.

          if ls_stylemapping-complete_stylex-font-italic = abap_true.
            ld_flag_italic = ls_stylemapping-complete_style-font-italic.
          endif.

        catch zcx_excel.                                "#EC NO_HANDLER
          " Style GUID is present, but style was not found
          " Continue with default values

      endtry.
    endif.

    ep_width = zcl_excel_font=>calculate_text_width(
      iv_font_name   = ld_font_name
      iv_font_height = ld_font_height
      iv_flag_bold   = ld_flag_bold
      iv_flag_italic = ld_flag_italic
      iv_cell_value  = ld_cell_value ).

    " If the current cell contains an auto filter, make it a bit wider.
    " The size used by the auto filter button does not depend on the font
    " size.
    if ld_flag_contains_auto_filter = abap_true.
      ep_width += 2.
    endif.

  endmethod.


  method calculate_column_widths.
    types:
      begin of t_auto_size,
        col_index type int4,
        width     type f,
      end   of t_auto_size.
    types: tt_auto_size type table of t_auto_size.

    data: lo_column_iterator type ref to zcl_excel_collection_iterator,
          lo_column          type ref to zcl_excel_column.

    data: auto_size   type abap_boolean.
    data: auto_sizes  type tt_auto_size.
    data: count       type int4.
    data: highest_row type int4.
    data: width       type f.

    field-symbols: <auto_size>        like line of auto_sizes.

    lo_column_iterator = me->get_columns_iterator( ).
    while lo_column_iterator->has_next( ) = abap_true.
      lo_column ?= lo_column_iterator->get_next( ).
      auto_size = lo_column->get_auto_size( ).
      if auto_size = abap_true.
        append initial line to auto_sizes assigning <auto_size>.
        <auto_size>-col_index = lo_column->get_column_index( ).
        <auto_size>-width     = -1.
      endif.
    endwhile.

    " There is only something to do if there are some auto-size columns
    if not auto_sizes is initial.
      highest_row = me->get_highest_row( ).
      loop at auto_sizes assigning <auto_size>.
        count = 1.
        while count <= highest_row.
* Do not check merged cells
          if is_cell_merged(
              ip_column    = <auto_size>-col_index
              ip_row       = count ) = abap_false.
            width = calculate_cell_width( ip_column = <auto_size>-col_index     " issue #155 - less restrictive typing for ip_column
                                          ip_row    = count ).
            if width > <auto_size>-width.
              <auto_size>-width = width.
            endif.
          endif.
          count = count + 1.
        endwhile.
        lo_column = me->get_column( <auto_size>-col_index ). " issue #155 - less restrictive typing for ip_column
        lo_column->set_width( <auto_size>-width ).
      endloop.
    endif.

  endmethod.                    "CALCULATE_COLUMN_WIDTHS


  method calculate_table_bottom_right.

    data: lv_errormessage type string,
          lv_columns      type i,
          lt_columns      type zif_excel_data_decl=>zexcel_t_fieldcatalog,
          lv_maxrow       type i,
          lo_iterator     type ref to zcl_excel_collection_iterator,
          lo_curtable     type ref to zcl_excel_table,
          lv_row_int      type zif_excel_data_decl=>zexcel_cell_row,
          lv_column_int   type zif_excel_data_decl=>zexcel_cell_column,
          lv_rows         type i,
          lv_maxcol       type i.

    "Get the number of columns for the current table
    lt_columns = it_field_catalog.
    lv_columns = lines( lt_columns ).

    "Calculate the top left row of the current table
    lv_column_int = zcl_excel_common=>convert_column2int( cs_settings-top_left_column ).
    lv_row_int    = cs_settings-top_left_row.

    "Get number of row for the current table
    lv_rows = lines( ip_table ).

    "Calculate the bottom right row for the current table
    lv_maxcol                       = lv_column_int + lv_columns - 1.
    lv_maxrow                       = lv_row_int    + lv_rows.
    cs_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( lv_maxcol ).
    cs_settings-bottom_right_row    = lv_maxrow.

  endmethod.


  method change_area_style.

    data: lv_row              type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_start        type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_to           type zif_excel_data_decl=>zexcel_cell_row,
          lv_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          lv_column_start_int type zif_excel_data_decl=>zexcel_cell_column,
          lv_column_end_int   type zif_excel_data_decl=>zexcel_cell_column.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               importing ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_to ).

    lv_column_int = lv_column_start_int.
    while lv_column_int <= lv_column_end_int.

      lv_row = lv_row_start.
      while lv_row <= lv_row_to.

        ip_style_changer->apply( ip_worksheet = me
                                 ip_column    = lv_column_int
                                 ip_row       = lv_row ).

        lv_row += 1.
      endwhile.

      lv_column_int += 1.
    endwhile.

  endmethod.


  method change_cell_style.

    data: changer type ref to zif_excel_style_changer,
          column  type zif_excel_data_decl=>zexcel_cell_column,
          row     type zif_excel_data_decl=>zexcel_cell_row.

    normalize_columnrow_parameter( exporting ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   importing ep_column    = column
                                             ep_row       = row ).

    changer = zcl_excel_style_changer=>create( excel = excel ).


    if ip_complete is supplied.
      if ip_xcomplete is not supplied.
        zcx_excel=>raise_text( 'Complete styleinfo has to be supplied with corresponding X-field' ).
      endif.
      changer->set_complete( ip_complete = ip_complete ip_xcomplete = ip_xcomplete ).
    endif.



    if ip_font is supplied.
      if ip_xfont is supplied.
        changer->set_complete_font( ip_font = ip_font ip_xfont = ip_xfont ).
      else.
        changer->set_complete_font( ip_font = ip_font ).
      endif.
    endif.

    if ip_fill is supplied.
      if ip_xfill is supplied.
        changer->set_complete_fill( ip_fill = ip_fill ip_xfill = ip_xfill ).
      else.
        changer->set_complete_fill( ip_fill = ip_fill ).
      endif.
    endif.


    if ip_borders is supplied.
      if ip_xborders is supplied.
        changer->set_complete_borders( ip_borders = ip_borders ip_xborders = ip_xborders ).
      else.
        changer->set_complete_borders( ip_borders = ip_borders ).
      endif.
    endif.

    if ip_alignment is supplied.
      if ip_xalignment is supplied.
        changer->set_complete_alignment( ip_alignment = ip_alignment ip_xalignment = ip_xalignment ).
      else.
        changer->set_complete_alignment( ip_alignment = ip_alignment ).
      endif.
    endif.

    if ip_protection is supplied.
      if ip_xprotection is supplied.
        changer->set_complete_protection( ip_protection = ip_protection ip_xprotection = ip_xprotection ).
      else.
        changer->set_complete_protection( ip_protection = ip_protection ).
      endif.
    endif.


    if ip_borders_allborders is supplied.
      if ip_xborders_allborders is supplied.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ip_xborders_allborders = ip_xborders_allborders ).
      else.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ).
      endif.
    endif.

    if ip_borders_diagonal is supplied.
      if ip_xborders_diagonal is supplied.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ip_xborders_diagonal = ip_xborders_diagonal ).
      else.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ).
      endif.
    endif.

    if ip_borders_down is supplied.
      if ip_xborders_down is supplied.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ip_xborders_down = ip_xborders_down ).
      else.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ).
      endif.
    endif.

    if ip_borders_left is supplied.
      if ip_xborders_left is supplied.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ip_xborders_left = ip_xborders_left ).
      else.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ).
      endif.
    endif.

    if ip_borders_right is supplied.
      if ip_xborders_right is supplied.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ip_xborders_right = ip_xborders_right ).
      else.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ).
      endif.
    endif.

    if ip_borders_top is supplied.
      if ip_xborders_top is supplied.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ip_xborders_top = ip_xborders_top ).
      else.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ).
      endif.
    endif.

    if ip_number_format_format_code is supplied.
      changer->set_number_format( ip_number_format_format_code ).
    endif.
    if ip_font_bold is supplied.
      changer->set_font_bold( ip_font_bold ).
    endif.
    if ip_font_color is supplied.
      changer->set_font_color( ip_font_color ).
    endif.
    if ip_font_color_rgb is supplied.
      changer->set_font_color_rgb( ip_font_color_rgb ).
    endif.
    if ip_font_color_indexed is supplied.
      changer->set_font_color_indexed( ip_font_color_indexed ).
    endif.
    if ip_font_color_theme is supplied.
      changer->set_font_color_theme( ip_font_color_theme ).
    endif.
    if ip_font_color_tint is supplied.
      changer->set_font_color_tint( ip_font_color_tint ).
    endif.

    if ip_font_family is supplied.
      changer->set_font_family( ip_font_family ).
    endif.
    if ip_font_italic is supplied.
      changer->set_font_italic( ip_font_italic ).
    endif.
    if ip_font_name is supplied.
      changer->set_font_name( ip_font_name ).
    endif.
    if ip_font_scheme is supplied.
      changer->set_font_scheme( ip_font_scheme ).
    endif.
    if ip_font_size is supplied.
      changer->set_font_size( ip_font_size ).
    endif.
    if ip_font_strikethrough is supplied.
      changer->set_font_strikethrough( ip_font_strikethrough ).
    endif.
    if ip_font_underline is supplied.
      changer->set_font_underline( ip_font_underline ).
    endif.
    if ip_font_underline_mode is supplied.
      changer->set_font_underline_mode( ip_font_underline_mode ).
    endif.

    if ip_fill_filltype is supplied.
      changer->set_fill_filltype( ip_fill_filltype ).
    endif.
    if ip_fill_rotation is supplied.
      changer->set_fill_rotation( ip_fill_rotation ).
    endif.
    if ip_fill_fgcolor is supplied.
      changer->set_fill_fgcolor( ip_fill_fgcolor ).
    endif.
    if ip_fill_fgcolor_rgb is supplied.
      changer->set_fill_fgcolor_rgb( ip_fill_fgcolor_rgb ).
    endif.
    if ip_fill_fgcolor_indexed is supplied.
      changer->set_fill_fgcolor_indexed( ip_fill_fgcolor_indexed ).
    endif.
    if ip_fill_fgcolor_theme is supplied.
      changer->set_fill_fgcolor_theme( ip_fill_fgcolor_theme ).
    endif.
    if ip_fill_fgcolor_tint is supplied.
      changer->set_fill_fgcolor_tint( ip_fill_fgcolor_tint ).
    endif.

    if ip_fill_bgcolor is supplied.
      changer->set_fill_bgcolor( ip_fill_bgcolor ).
    endif.
    if ip_fill_bgcolor_rgb is supplied.
      changer->set_fill_bgcolor_rgb( ip_fill_bgcolor_rgb ).
    endif.
    if ip_fill_bgcolor_indexed is supplied.
      changer->set_fill_bgcolor_indexed( ip_fill_bgcolor_indexed ).
    endif.
    if ip_fill_bgcolor_theme is supplied.
      changer->set_fill_bgcolor_theme( ip_fill_bgcolor_theme ).
    endif.
    if ip_fill_bgcolor_tint is supplied.
      changer->set_fill_bgcolor_tint( ip_fill_bgcolor_tint ).
    endif.

    if ip_fill_gradtype_type is supplied.
      changer->set_fill_gradtype_type( ip_fill_gradtype_type ).
    endif.
    if ip_fill_gradtype_degree is supplied.
      changer->set_fill_gradtype_degree( ip_fill_gradtype_degree ).
    endif.
    if ip_fill_gradtype_bottom is supplied.
      changer->set_fill_gradtype_bottom( ip_fill_gradtype_bottom ).
    endif.
    if ip_fill_gradtype_left is supplied.
      changer->set_fill_gradtype_left( ip_fill_gradtype_left ).
    endif.
    if ip_fill_gradtype_top is supplied.
      changer->set_fill_gradtype_top( ip_fill_gradtype_top ).
    endif.
    if ip_fill_gradtype_right is supplied.
      changer->set_fill_gradtype_right( ip_fill_gradtype_right ).
    endif.
    if ip_fill_gradtype_position1 is supplied.
      changer->set_fill_gradtype_position1( ip_fill_gradtype_position1 ).
    endif.
    if ip_fill_gradtype_position2 is supplied.
      changer->set_fill_gradtype_position2( ip_fill_gradtype_position2 ).
    endif.
    if ip_fill_gradtype_position3 is supplied.
      changer->set_fill_gradtype_position3( ip_fill_gradtype_position3 ).
    endif.



    if ip_borders_diagonal_mode is supplied.
      changer->set_borders_diagonal_mode( ip_borders_diagonal_mode ).
    endif.
    if ip_alignment_horizontal is supplied.
      changer->set_alignment_horizontal( ip_alignment_horizontal ).
    endif.
    if ip_alignment_vertical is supplied.
      changer->set_alignment_vertical( ip_alignment_vertical ).
    endif.
    if ip_alignment_textrotation is supplied.
      changer->set_alignment_textrotation( ip_alignment_textrotation ).
    endif.
    if ip_alignment_wraptext is supplied.
      changer->set_alignment_wraptext( ip_alignment_wraptext ).
    endif.
    if ip_alignment_shrinktofit is supplied.
      changer->set_alignment_shrinktofit( ip_alignment_shrinktofit ).
    endif.
    if ip_alignment_indent is supplied.
      changer->set_alignment_indent( ip_alignment_indent ).
    endif.
    if ip_protection_hidden is supplied.
      changer->set_protection_hidden( ip_protection_hidden ).
    endif.
    if ip_protection_locked is supplied.
      changer->set_protection_locked( ip_protection_locked ).
    endif.

    if ip_borders_allborders_style is supplied.
      changer->set_borders_allborders_style( ip_borders_allborders_style ).
    endif.
    if ip_borders_allborders_color is supplied.
      changer->set_borders_allborders_color( ip_borders_allborders_color ).
    endif.
    if ip_borders_allbo_color_rgb is supplied.
      changer->set_borders_allbo_color_rgb( ip_borders_allbo_color_rgb ).
    endif.
    if ip_borders_allbo_color_indexed is supplied.
      changer->set_borders_allbo_color_indexe( ip_borders_allbo_color_indexed ).
    endif.
    if ip_borders_allbo_color_theme is supplied.
      changer->set_borders_allbo_color_theme( ip_borders_allbo_color_theme ).
    endif.
    if ip_borders_allbo_color_tint is supplied.
      changer->set_borders_allbo_color_tint( ip_borders_allbo_color_tint ).
    endif.

    if ip_borders_diagonal_style is supplied.
      changer->set_borders_diagonal_style( ip_borders_diagonal_style ).
    endif.
    if ip_borders_diagonal_color is supplied.
      changer->set_borders_diagonal_color( ip_borders_diagonal_color ).
    endif.
    if ip_borders_diagonal_color_rgb is supplied.
      changer->set_borders_diagonal_color_rgb( ip_borders_diagonal_color_rgb ).
    endif.
    if ip_borders_diagonal_color_inde is supplied.
      changer->set_borders_diagonal_color_ind( ip_borders_diagonal_color_inde ).
    endif.
    if ip_borders_diagonal_color_them is supplied.
      changer->set_borders_diagonal_color_the( ip_borders_diagonal_color_them ).
    endif.
    if ip_borders_diagonal_color_tint is supplied.
      changer->set_borders_diagonal_color_tin( ip_borders_diagonal_color_tint ).
    endif.

    if ip_borders_down_style is supplied.
      changer->set_borders_down_style( ip_borders_down_style ).
    endif.
    if ip_borders_down_color is supplied.
      changer->set_borders_down_color( ip_borders_down_color ).
    endif.
    if ip_borders_down_color_rgb is supplied.
      changer->set_borders_down_color_rgb( ip_borders_down_color_rgb ).
    endif.
    if ip_borders_down_color_indexed is supplied.
      changer->set_borders_down_color_indexed( ip_borders_down_color_indexed ).
    endif.
    if ip_borders_down_color_theme is supplied.
      changer->set_borders_down_color_theme( ip_borders_down_color_theme ).
    endif.
    if ip_borders_down_color_tint is supplied.
      changer->set_borders_down_color_tint( ip_borders_down_color_tint ).
    endif.

    if ip_borders_left_style is supplied.
      changer->set_borders_left_style( ip_borders_left_style ).
    endif.
    if ip_borders_left_color is supplied.
      changer->set_borders_left_color( ip_borders_left_color ).
    endif.
    if ip_borders_left_color_rgb is supplied.
      changer->set_borders_left_color_rgb( ip_borders_left_color_rgb ).
    endif.
    if ip_borders_left_color_indexed is supplied.
      changer->set_borders_left_color_indexed( ip_borders_left_color_indexed ).
    endif.
    if ip_borders_left_color_theme is supplied.
      changer->set_borders_left_color_theme( ip_borders_left_color_theme ).
    endif.
    if ip_borders_left_color_tint is supplied.
      changer->set_borders_left_color_tint( ip_borders_left_color_tint ).
    endif.

    if ip_borders_right_style is supplied.
      changer->set_borders_right_style( ip_borders_right_style ).
    endif.
    if ip_borders_right_color is supplied.
      changer->set_borders_right_color( ip_borders_right_color ).
    endif.
    if ip_borders_right_color_rgb is supplied.
      changer->set_borders_right_color_rgb( ip_borders_right_color_rgb ).
    endif.
    if ip_borders_right_color_indexed is supplied.
      changer->set_borders_right_color_indexe( ip_borders_right_color_indexed ).
    endif.
    if ip_borders_right_color_theme is supplied.
      changer->set_borders_right_color_theme( ip_borders_right_color_theme ).
    endif.
    if ip_borders_right_color_tint is supplied.
      changer->set_borders_right_color_tint( ip_borders_right_color_tint ).
    endif.

    if ip_borders_top_style is supplied.
      changer->set_borders_top_style( ip_borders_top_style ).
    endif.
    if ip_borders_top_color is supplied.
      changer->set_borders_top_color( ip_borders_top_color ).
    endif.
    if ip_borders_top_color_rgb is supplied.
      changer->set_borders_top_color_rgb( ip_borders_top_color_rgb ).
    endif.
    if ip_borders_top_color_indexed is supplied.
      changer->set_borders_top_color_indexed( ip_borders_top_color_indexed ).
    endif.
    if ip_borders_top_color_theme is supplied.
      changer->set_borders_top_color_theme( ip_borders_top_color_theme ).
    endif.
    if ip_borders_top_color_tint is supplied.
      changer->set_borders_top_color_tint( ip_borders_top_color_tint ).
    endif.


    ep_guid = changer->apply( ip_worksheet = me
                              ip_column    = column
                              ip_row       = row ).


  endmethod.                    "CHANGE_CELL_STYLE


  method check_cell_column_formula.

    field-symbols <fs_column_formula> type zcl_excel_worksheet=>mty_s_column_formula.

    if ip_value is not initial or ip_formula is not initial.
      zcx_excel=>raise_text( c_messages-formula_id_only_is_possible ).
    endif.
    read table it_column_formulas with table key id = ip_column_formula_id assigning <fs_column_formula>.
    if sy-subrc <> 0.
      zcx_excel=>raise_text( c_messages-column_formula_id_not_found ).
    endif.
    if ip_row < <fs_column_formula>-table_top_left_row + 1
          or ip_row > <fs_column_formula>-table_bottom_right_row + 1
          or ip_column < <fs_column_formula>-table_left_column_int
          or ip_column > <fs_column_formula>-table_right_column_int.
      zcx_excel=>raise_text( c_messages-formula_not_in_this_table ).
    endif.
    if ip_column <> <fs_column_formula>-column.
      zcx_excel=>raise_text( c_messages-formula_in_other_column ).
    endif.

  endmethod.


  method check_rtf.

    data: lo_style           type ref to zcl_excel_style,
          ls_font            type zif_excel_data_decl=>zexcel_s_style_font,
          lv_next_rtf_offset type i,
          lv_tabix           type i,
          lv_value           type string,
          lv_val_length      type i,
          ls_rtf             like line of ct_rtf.
    field-symbols: <rtf> like line of ct_rtf.

    if ip_style is not supplied.
      ip_style = excel->get_default_style( ).
    endif.

    lo_style = excel->get_style_from_guid( ip_style ).
    if lo_style is bound.
      ls_font  = lo_style->font->get_structure( ).
    endif.

    lv_next_rtf_offset = 0.
    loop at ct_rtf assigning <rtf>.
      lv_tabix = sy-tabix.
      if lv_next_rtf_offset < <rtf>-offset.
        ls_rtf-offset = lv_next_rtf_offset.
        ls_rtf-length = <rtf>-offset - lv_next_rtf_offset.
        ls_rtf-font   = ls_font.
        insert ls_rtf into ct_rtf index lv_tabix.
      elseif lv_next_rtf_offset > <rtf>-offset.
        raise exception type zcx_excel
          exporting
            error = 'Gaps or overlaps in RTF data offset/length specs'.
      endif.
      lv_next_rtf_offset = <rtf>-offset + <rtf>-length.
    endloop.

    lv_value = ip_value.
    lv_val_length = strlen( lv_value ).
    if lv_val_length > lv_next_rtf_offset.
      ls_rtf-offset = lv_next_rtf_offset.
      ls_rtf-length = lv_val_length - lv_next_rtf_offset.
      ls_rtf-font   = ls_font.
      insert ls_rtf into table ct_rtf.
    elseif lv_val_length < lv_next_rtf_offset.
      raise exception type zcx_excel
        exporting
          error = 'RTF specs length is not equal to value length'.
    endif.

  endmethod.


  method check_table_overlapping.

    data: lv_errormessage type string,
          lv_column_int   type zif_excel_data_decl=>zexcel_cell_column,
          lv_maxcol       type i.
    field-symbols:
          <ls_table_settings> type zif_excel_data_decl=>zexcel_s_table_settings.

    lv_column_int = zcl_excel_common=>convert_column2int( is_table_settings-top_left_column ).
    lv_maxcol = zcl_excel_common=>convert_column2int( is_table_settings-bottom_right_column ).

    loop at it_other_table_settings assigning <ls_table_settings>.

      if  (    (  is_table_settings-top_left_row     ge <ls_table_settings>-top_left_row
              and is_table_settings-top_left_row     le <ls_table_settings>-bottom_right_row )
            or
               (  is_table_settings-bottom_right_row ge <ls_table_settings>-top_left_row
              and is_table_settings-bottom_right_row le <ls_table_settings>-bottom_right_row )
          )
        and
          (    (  lv_column_int ge zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              and lv_column_int le zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
            or
               (  lv_maxcol     ge zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              and lv_maxcol     le zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
          ).
        lv_errormessage = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
        zcx_excel=>raise_text( lv_errormessage ).
      endif.

    endloop.

  endmethod.


  method class_constructor.
    field-symbols <lv_typekind> type abap_typekind.
    data lo_rtti type ref to cl_abap_datadescr.

    c_messages-formula_id_only_is_possible = |{ 'If Formula ID is used, value and formula must be empty'(008) }|.
    c_messages-column_formula_id_not_found = |{ 'The Column Formula does not exist'(009) }|.
    c_messages-formula_not_in_this_table = |{ 'The cell uses a Column Formula which should be part of the same table'(010) }|.
    c_messages-formula_in_other_column = |{ 'The cell uses a Column Formula which is in a different column'(011) }|.

    call method cl_abap_elemdescr=>get_utclong receiving p_result = lo_rtti.
    create data variable_utclong type handle lo_rtti.

  endmethod.


  method clear_initial_colorxfields.

    if is_color-rgb is initial.
      clear cs_xcolor-rgb.
    endif.
    if is_color-indexed is initial.
      clear cs_xcolor-indexed.
    endif.
    if is_color-theme is initial.
      clear cs_xcolor-theme.
    endif.
    if is_color-tint is initial.
      clear cs_xcolor-tint.
    endif.

  endmethod.


  method constructor.
    data: lv_title type zif_excel_data_decl=>zexcel_sheet_title.

    me->excel = ip_excel.

    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).        " ins issue #379 - replacement for outdated function call

    if ip_title is not initial.
      lv_title = ip_title.
    else.
      lv_title = me->generate_title( ). " ins issue #154 - Names of worksheets
    endif.

    me->set_title( ip_title = lv_title ).

    create object sheet_setup.
    create object styles_cond.
    create object data_validations.
    create object tables.
    create object columns.
    create object rows.
    create object ranges. " issue #163
    create object mo_pagebreaks.
    create object drawings
      exporting
        ip_type = zcl_excel_drawing=>type_image.
    create object charts
      exporting
        ip_type = zcl_excel_drawing=>type_chart.
    me->zif_excel_sheet_protection~initialize( ).
    me->zif_excel_sheet_properties~initialize( ).
    create object hyperlinks.
    create object comments. " (+) Issue #180

* initialize active cell coordinates
    active_cell-cell_row = 1.
    active_cell-cell_column = 1.

* inizialize dimension range
    lower_cell-cell_row     = 1.
    lower_cell-cell_column  = 1.
    upper_cell-cell_row     = 1.
    upper_cell-cell_column  = 1.

  endmethod.                    "CONSTRUCTOR


  method convert_to_table.
*@TODO:Commented by Juwin
*    TYPES:
*      BEGIN OF ts_field_conv,
*        fieldname TYPE x031l-fieldname,
*        convexit  TYPE x031l-convexit,
*      END OF ts_field_conv,
*      BEGIN OF ts_style_conv,
*        cell_style type zif_excel_data_decl=>zexcel_s_cell_data-cell_style,
*        abap_type  TYPE abap_typekind,
*      END OF ts_style_conv.
*
*    DATA:
*      lv_row_int          type zif_excel_data_decl=>zexcel_cell_row,
*      lv_column_int       type zif_excel_data_decl=>zexcel_cell_column,
*      lv_column_alpha     type zif_excel_data_decl=>zexcel_cell_column_alpha,
*      lt_field_catalog    type zif_excel_data_decl=>zexcel_t_fieldcatalog,
*      ls_field_catalog    type zif_excel_data_decl=>zexcel_s_fieldcatalog,
*      lv_value            TYPE string,
*      lv_maxcol           TYPE i,
*      lv_maxrow           TYPE i,
*      lt_field_conv       TYPE TABLE OF ts_field_conv,
*      lt_comp             TYPE abap_component_tab,
*      ls_comp             TYPE abap_componentdescr,
*      lo_line_type        TYPE REF TO cl_abap_structdescr,
*      lo_tab_type         TYPE REF TO cl_abap_tabledescr,
*      lr_data             TYPE REF TO data,
*      lt_comp_view        TYPE abap_component_view_tab,
*      ls_comp_view        TYPE abap_simple_componentdescr,
*      lt_ddic_object      TYPE dd_x031l_table,
*      lt_ddic_object_comp TYPE dd_x031l_table,
*      ls_ddic_object      TYPE x031l,
*      lt_style_conv       TYPE TABLE OF ts_style_conv,
*      ls_style_conv       TYPE ts_style_conv,
*      ls_stylemapping     type zif_excel_data_decl=>zexcel_s_stylemapping,
*      lv_format_code      type zif_excel_data_decl=>zexcel_number_format,
*      lv_float            TYPE f,
*      lt_map_excel_row    TYPE TABLE OF i,
*      lv_index            TYPE i,
*      lv_index_col        TYPE i.
*
*    FIELD-SYMBOLS:
*      <lt_data>          TYPE STANDARD TABLE,
*      <ls_data>          TYPE data,
*      <lv_data>          TYPE data,
*      <lt_data2>         TYPE STANDARD TABLE,
*      <ls_data2>         TYPE data,
*      <lv_data2>         TYPE data,
*      <ls_field_conv>    TYPE ts_field_conv,
*      <ls_ddic_object>   TYPE x031l,
*      <ls_sheet_content> type zif_excel_data_decl=>zexcel_s_cell_data,
*      <fs_typekind_int8> TYPE abap_typekind.
*
*    CLEAR: et_data, er_data.
*
*    lv_maxcol = get_highest_column( ).
*    lv_maxrow = get_highest_row( ).
*
*
*    " Field catalog
*    lt_field_catalog = it_field_catalog.
*    IF lt_field_catalog IS INITIAL.
*      IF et_data IS SUPPLIED.
*        lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = et_data ).
*      ELSE.
*        DO lv_maxcol TIMES.
*          ls_field_catalog-position = sy-index.
*          ls_field_catalog-fieldname = 'COL_' && sy-index.
*          ls_field_catalog-dynpfld = abap_true.
*          APPEND ls_field_catalog TO lt_field_catalog.
*        ENDDO.
*      ENDIF.
*    ENDIF.
*
*    SORT lt_field_catalog BY position.
*    DELETE lt_field_catalog WHERE dynpfld NE abap_true.
*    CHECK: lt_field_catalog IS NOT INITIAL.
*
*
*    " Create dynamic table string columns
*    ls_comp-type = cl_abap_elemdescr=>get_string( ).
*    LOOP AT lt_field_catalog INTO ls_field_catalog.
*      ls_comp-name = ls_field_catalog-fieldname.
*      APPEND ls_comp TO lt_comp.
*    ENDLOOP.
*    lo_line_type = cl_abap_structdescr=>create( lt_comp ).
*    lo_tab_type = cl_abap_tabledescr=>create( lo_line_type ).
*    CREATE DATA er_data TYPE HANDLE lo_tab_type.
*    ASSIGN er_data->* TO <lt_data>.
*
*
*    " Collect field conversion rules
*    IF et_data IS SUPPLIED.
**      lt_ddic_object = get_ddic_object( et_data ).
*      lo_tab_type ?= cl_abap_tabledescr=>describe_by_data( et_data ).
*      lo_line_type ?= lo_tab_type->get_table_line_type( ).
*      lo_line_type->get_ddic_object(
*        RECEIVING
*          p_object = lt_ddic_object
*        EXCEPTIONS
*          OTHERS   = 3
*      ).
*      IF lt_ddic_object IS INITIAL.
*        lt_comp_view = lo_line_type->get_included_view( ).
*        LOOP AT lt_comp_view INTO ls_comp_view.
*          ls_comp_view-type->get_ddic_object(
*            RECEIVING
*              p_object = lt_ddic_object_comp
*            EXCEPTIONS
*              OTHERS   = 3
*          ).
*          IF lt_ddic_object_comp IS NOT INITIAL.
*            READ TABLE lt_ddic_object_comp INTO ls_ddic_object INDEX 1.
*            ls_ddic_object-fieldname = ls_comp_view-name.
*            APPEND ls_ddic_object TO lt_ddic_object.
*          ENDIF.
*        ENDLOOP.
*      ENDIF.
*
*      SORT lt_ddic_object BY fieldname.
*      LOOP AT lt_field_catalog INTO ls_field_catalog.
*        APPEND INITIAL LINE TO lt_field_conv ASSIGNING <ls_field_conv>.
*        MOVE-CORRESPONDING ls_field_catalog TO <ls_field_conv>.
*        READ TABLE lt_ddic_object ASSIGNING <ls_ddic_object> WITH KEY fieldname = <ls_field_conv>-fieldname BINARY SEARCH.
*        CHECK: sy-subrc EQ 0.
*
*        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
*        IF sy-subrc <> 0.
*          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
*        ENDIF.
*
*        CASE <ls_ddic_object>-exid.
*          WHEN cl_abap_typedescr=>typekind_int
*            OR cl_abap_typedescr=>typekind_int1
*            OR <fs_typekind_int8>
*            OR cl_abap_typedescr=>typekind_int2
*            OR cl_abap_typedescr=>typekind_packed
*            OR cl_abap_typedescr=>typekind_decfloat
*            OR cl_abap_typedescr=>typekind_decfloat16
*            OR cl_abap_typedescr=>typekind_decfloat34
*            OR cl_abap_typedescr=>typekind_float.
*            " Numbers
*            <ls_field_conv>-convexit = cl_abap_typedescr=>typekind_float.
*          WHEN OTHERS.
*            <ls_field_conv>-convexit = <ls_ddic_object>-convexit.
*        ENDCASE.
*      ENDLOOP.
*    ENDIF.
*
*    " Date & Time in excel style
*    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> WHERE cell_style IS NOT INITIAL AND data_type IS INITIAL. "#EC CI_SORTSEQ
*      ls_style_conv-cell_style = <ls_sheet_content>-cell_style.
*      APPEND ls_style_conv TO lt_style_conv.
*    ENDLOOP.
*    IF lt_style_conv IS NOT INITIAL.
*      SORT lt_style_conv BY cell_style.
*      DELETE ADJACENT DUPLICATES FROM lt_style_conv COMPARING cell_style.
*
*      LOOP AT lt_style_conv INTO ls_style_conv.
*
*        ls_stylemapping = me->excel->get_style_to_guid( ls_style_conv-cell_style ).
*        lv_format_code = ls_stylemapping-complete_style-number_format-format_code.
*        " https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
*        IF lv_format_code CS ';'.
*          lv_format_code = lv_format_code(sy-fdpos).
*        ENDIF.
*        CHECK: lv_format_code NA '#?'.
*
*        " Remove color pattern
*        REPLACE ALL OCCURRENCES OF REGEX '\[\L[^]]*\]' IN lv_format_code WITH ''.
*
*        IF lv_format_code CA 'yd' OR lv_format_code EQ zcl_excel_style_number_format=>c_format_date_std.
*          " DATE = yyyymmdd
*          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_date.
*        ELSEIF lv_format_code CA 'hs'.
*          " TIME = hhmmss
*          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_time.
*        ELSE.
*          DELETE lt_style_conv.
*          CONTINUE.
*        ENDIF.
*
*        MODIFY lt_style_conv FROM ls_style_conv TRANSPORTING abap_type.
*
*      ENDLOOP.
*    ENDIF.
*
*
**--------------------------------------------------------------------*
** Start of convert content
**--------------------------------------------------------------------*
*    READ TABLE me->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = iv_begin_row.
*    IF sy-subrc EQ 0.
*      lv_index = sy-tabix.
*    ENDIF.
*
*    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> FROM lv_index.
*      AT NEW cell_row.
*        IF iv_end_row <> 0
*        AND <ls_sheet_content>-cell_row > iv_end_row.
*          EXIT.
*        ENDIF.
*        " New line
*        APPEND INITIAL LINE TO <lt_data> ASSIGNING <ls_data>.
*        lv_index = sy-tabix.
*      ENDAT.
*
*      IF <ls_sheet_content>-cell_value IS NOT INITIAL.
*        ASSIGN COMPONENT <ls_sheet_content>-cell_column OF STRUCTURE <ls_data> TO <lv_data>.
*        IF sy-subrc EQ 0.
*          " value
*          <lv_data> = <ls_sheet_content>-cell_value.
*
*          " field conversion
*          READ TABLE lt_field_conv ASSIGNING <ls_field_conv> INDEX <ls_sheet_content>-cell_column.
*          IF sy-subrc EQ 0 AND <ls_field_conv>-convexit IS NOT INITIAL.
*            CASE <ls_field_conv>-convexit.
*              WHEN cl_abap_typedescr=>typekind_float.
*                lv_float = zcl_excel_common=>excel_string_to_number( <ls_sheet_content>-cell_value ).
*                <lv_data> = |{ lv_float NUMBER = RAW }|.
*              WHEN 'ALPHA'.
*                CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
*                  EXPORTING
*                    input  = <ls_sheet_content>-cell_value
*                  IMPORTING
*                    output = <lv_data>.
*            ENDCASE.
*          ENDIF.
*
*          " style conversion
*          IF <ls_sheet_content>-cell_style IS NOT INITIAL.
*            READ TABLE lt_style_conv INTO ls_style_conv WITH KEY cell_style = <ls_sheet_content>-cell_style BINARY SEARCH.
*            IF sy-subrc EQ 0.
*              CASE ls_style_conv-abap_type.
*                WHEN cl_abap_typedescr=>typekind_date.
*                  <lv_data> = zcl_excel_common=>excel_string_to_date( ip_value = <ls_sheet_content>-cell_value
*                                                                      ip_exact = abap_true ).
*                WHEN cl_abap_typedescr=>typekind_time.
*                  <lv_data> = zcl_excel_common=>excel_string_to_time( <ls_sheet_content>-cell_value ).
*              ENDCASE.
*            ENDIF.
*          ENDIF.
*
*          " condense
*          CONDENSE <lv_data>.
*        ENDIF.
*      ENDIF.
*
*      AT END OF cell_row.
*        " Delete empty line
*        IF <ls_data> IS INITIAL.
*          DELETE <lt_data> INDEX lv_index.
*        ELSE.
*          APPEND <ls_sheet_content>-cell_row TO lt_map_excel_row.
*        ENDIF.
*      ENDAT.
*    ENDLOOP.
**--------------------------------------------------------------------*
** End of convert content
**--------------------------------------------------------------------*
*
*
*    IF et_data IS SUPPLIED.
**      MOVE-CORRESPONDING <lt_data> TO et_data.
*      LOOP AT <lt_data> ASSIGNING <ls_data>.
*        APPEND INITIAL LINE TO et_data ASSIGNING <ls_data2>.
*        MOVE-CORRESPONDING <ls_data> TO <ls_data2>.
*      ENDLOOP.
*    ENDIF.
*
*    " Apply conversion exit.
*    LOOP AT lt_field_conv ASSIGNING <ls_field_conv>
*     WHERE convexit = 'ALPHA'.
*      LOOP AT et_data ASSIGNING <ls_data>.
*        ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
*        CHECK: sy-subrc EQ 0 AND <lv_data> IS NOT INITIAL.
*        CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
*          EXPORTING
*            input  = <lv_data>
*          IMPORTING
*            output = <lv_data>.
*      ENDLOOP.
*    ENDLOOP.

  endmethod.


  method delete_merge.

    data: lv_column type i.
*--------------------------------------------------------------------*
* If cell information is passed delete merge including this cell,
* otherwise delete all merges
*--------------------------------------------------------------------*
    if   ip_cell_column is initial
      or ip_cell_row    is initial.
      clear me->mt_merged_cells.
    else.
      lv_column = zcl_excel_common=>convert_column2int( ip_cell_column ).

      loop at me->mt_merged_cells transporting no fields
      where row_from <= ip_cell_row and row_to >= ip_cell_row
        and col_from <= lv_column and col_to >= lv_column. "#EC CI_SORTSEQ
        delete me->mt_merged_cells.
        exit.
      endloop.
    endif.

  endmethod.                    "DELETE_MERGE


  method delete_row_outline.

    delete me->mt_row_outlines where row_from = iv_row_from
                                 and row_to   = iv_row_to.
    if sy-subrc <> 0.  " didn't find outline that was to be deleted
      zcx_excel=>raise_text( 'Row outline to be deleted does not exist' ).
    endif.

  endmethod.                    "DELETE_ROW_OUTLINE


  method freeze_panes.

    if ip_num_columns is not supplied and ip_num_rows is not supplied.
      zcx_excel=>raise_text( 'Pleas provide number of rows and/or columns to freeze' ).
    endif.

    if ip_num_columns is supplied and ip_num_columns <= 0.
      zcx_excel=>raise_text( 'Number of columns to freeze should be positive' ).
    endif.

    if ip_num_rows is supplied and ip_num_rows <= 0.
      zcx_excel=>raise_text( 'Number of rows to freeze should be positive' ).
    endif.

    freeze_pane_cell_column = ip_num_columns + 1.
    freeze_pane_cell_row = ip_num_rows + 1.
  endmethod.                    "FREEZE_PANES


  method generate_title.
    data: lo_worksheets_iterator type ref to zcl_excel_collection_iterator,
          lo_worksheet           type ref to zcl_excel_worksheet.

    data: t_titles    type hashed table of zif_excel_data_decl=>zexcel_sheet_title with unique key table_line,
          title       type zif_excel_data_decl=>zexcel_sheet_title,
          sheetnumber type i.

* Get list of currently used titles
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    while lo_worksheets_iterator->has_next( ) = abap_true.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      title = lo_worksheet->get_title( ).
      insert title into table t_titles.
      sheetnumber += 1.
    endwhile.

* Now build sheetnumber.  Increase counter until we hit a number that is not used so far
    sheetnumber += 1.  " Start counting with next number
    do.
      title = sheetnumber.
      shift title left deleting leading space.
      concatenate 'Sheet'(001) title into ep_title.
      insert ep_title into table t_titles.
      if sy-subrc = 0.  " Title not used so far --> take it
        exit.
      endif.

      sheetnumber += 1.
    enddo.
  endmethod.                    "GENERATE_TITLE


  method get_active_cell.

    data: lv_active_column type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_active_row    type string.

    lv_active_column = zcl_excel_common=>convert_column2alpha( active_cell-cell_column ).
    lv_active_row    = active_cell-cell_row.
    shift lv_active_row right deleting trailing space.
    shift lv_active_row left deleting leading space.
    concatenate lv_active_column lv_active_row into ep_active_cell.

  endmethod.                    "GET_ACTIVE_CELL


  method get_cell.

    data: lv_column        type zif_excel_data_decl=>zexcel_cell_column,
          lv_row           type zif_excel_data_decl=>zexcel_cell_row,
          ls_sheet_content type zif_excel_data_decl=>zexcel_s_cell_data.

    normalize_columnrow_parameter( exporting ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   importing ep_column    = lv_column
                                             ep_row       = lv_row ).

    read table sheet_content into ls_sheet_content with table key cell_row     = lv_row
                                                                  cell_column  = lv_column.

    ep_rc       = sy-subrc.
    ep_value    = ls_sheet_content-cell_value.
    ep_guid     = ls_sheet_content-cell_style.       " issue 139 - added this to be used for columnwidth calculation
    ep_formula  = ls_sheet_content-cell_formula.
    if et_rtf is supplied and ls_sheet_content-rtf_tab is not initial.
      et_rtf = ls_sheet_content-rtf_tab.
    endif.

    " Addition to solve issue #120, contribution by Stefan Schmöcker
    data: style_iterator type ref to zcl_excel_collection_iterator,
          style          type ref to zcl_excel_style.
    if ep_style is supplied.
      clear ep_style.
      style_iterator = me->excel->get_styles_iterator( ).
      while style_iterator->has_next( ) = abap_true.
        style ?= style_iterator->get_next( ).
        if style->get_guid( ) = ls_sheet_content-cell_style.
          ep_style = style.
          exit.
        endif.
      endwhile.
    endif.
  endmethod.                    "GET_CELL


  method get_column.

    data: lv_column type zif_excel_data_decl=>zexcel_cell_column.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    eo_column = me->columns->get( ip_index = lv_column ).

    if eo_column is not bound.
      eo_column = me->add_new_column( ip_column ).
    endif.

  endmethod.                    "GET_COLUMN


  method get_columns.

    data: columns type table of i,
          column  type i.
    field-symbols:
          <sheet_cell> type zif_excel_data_decl=>zexcel_s_cell_data.

    loop at sheet_content assigning <sheet_cell>.
      collect <sheet_cell>-cell_column into columns.
    endloop.

    loop at columns into column.
      " This will create the column instance if it doesn't exist
      get_column( column ).
    endloop.

    eo_columns = me->columns.
  endmethod.                    "GET_COLUMNS


  method get_columns_iterator.

    get_columns( ).
    eo_iterator = me->columns->get_iterator( ).

  endmethod.                    "GET_COLUMNS_ITERATOR


  method get_comments.

    if iv_copy_collection = abap_true.
* By default, get_comments copies the collection (backward compatibility)
      create object r_comments
        exporting
          io_from = comments.
    else.
      r_comments = comments.
    endif.

  endmethod.                    "get_comments


  method get_comments_iterator.
    eo_iterator = comments->get_iterator( ).

  endmethod.                    "get_comments_iterator


  method get_data_validations_iterator.

    eo_iterator = me->data_validations->get_iterator( ).
  endmethod.                    "GET_DATA_VALIDATIONS_ITERATOR


  method get_data_validations_size.
    ep_size = me->data_validations->size( ).
  endmethod.                    "GET_DATA_VALIDATIONS_SIZE


  method get_default_column.
    if me->column_default is not bound.
      create object me->column_default
        exporting
          ip_index     = 'A'         " ????
          ip_worksheet = me
          ip_excel     = me->excel.
    endif.

    eo_column = me->column_default.
  endmethod.                    "GET_DEFAULT_COLUMN


  method get_default_excel_date_format.
    constants c_lang_e type sylangu value 'E'.

    if default_excel_date_format is not initial.
      ep_default_excel_date_format = default_excel_date_format.
      return.
    endif.

    "try to get defaults
    try.
        cl_abap_datfm=>get_date_format_des( exporting im_langu = c_lang_e
                                            importing ex_dateformat = default_excel_date_format ).
      catch cx_abap_datfm_format_unknown.

    endtry.

    " and fallback to fixed format
    if default_excel_date_format is initial.
      default_excel_date_format = zcl_excel_style_number_format=>c_format_date_ddmmyyyydot.
    endif.

    ep_default_excel_date_format = default_excel_date_format.
  endmethod.                    "GET_DEFAULT_EXCEL_DATE_FORMAT


  method get_default_excel_time_format.
    data: l_timefm type xutimefm.

    if default_excel_time_format is not initial.
      ep_default_excel_time_format = default_excel_time_format.
      return.
    endif.

* Let's get default
    l_timefm = cl_abap_timefm=>get_environment_timefm( ).
    case l_timefm.
      when 0.
*0  24 Hour Format (Example: 12:05:10)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
      when 1.
*1  12 Hour Format (Example: 12:05:10 PM)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      when 2.
*2  12 Hour Format (Example: 12:05:10 pm) for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      when 3.
*3  Hours from 0 to 11 (Example: 00:05:10 PM)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      when 4.
*4  Hours from 0 to 11 (Example: 00:05:10 pm)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      when others.
        " and fallback to fixed format
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
    endcase.

    ep_default_excel_time_format = default_excel_time_format.
  endmethod.                    "GET_DEFAULT_EXCEL_TIME_FORMAT


  method get_default_row.
    if me->row_default is not bound.
      create object me->row_default.
    endif.

    eo_row = me->row_default.
  endmethod.                    "GET_DEFAULT_ROW


  method get_dimension_range.

    me->update_dimension_range( ).
    if upper_cell eq lower_cell. "only one cell
      " Worksheet not filled
      if upper_cell-cell_coords is initial.
        ep_dimension_range = 'A1'.
      else.
        ep_dimension_range = upper_cell-cell_coords.
      endif.
    else.
      concatenate upper_cell-cell_coords ':' lower_cell-cell_coords into ep_dimension_range.
    endif.

  endmethod.                    "GET_DIMENSION_RANGE


  method get_drawings.

    data: lo_drawing  type ref to zcl_excel_drawing,
          lo_iterator type ref to zcl_excel_collection_iterator.

    case ip_type.
      when zcl_excel_drawing=>type_image.
        r_drawings = drawings.
      when zcl_excel_drawing=>type_chart.
        r_drawings = charts.
      when space.
        create object r_drawings
          exporting
            ip_type = ''.

        lo_iterator = drawings->get_iterator( ).
        while lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        endwhile.
        lo_iterator = charts->get_iterator( ).
        while lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        endwhile.
      when others.
    endcase.
  endmethod.                    "GET_DRAWINGS


  method get_drawings_iterator.
    case ip_type.
      when zcl_excel_drawing=>type_image.
        eo_iterator = drawings->get_iterator( ).
      when zcl_excel_drawing=>type_chart.
        eo_iterator = charts->get_iterator( ).
    endcase.
  endmethod.                    "GET_DRAWINGS_ITERATOR


  method get_freeze_cell.
    ep_row = me->freeze_pane_cell_row.
    ep_column = me->freeze_pane_cell_column.
  endmethod.                    "GET_FREEZE_CELL


  method get_guid.

    ep_guid = me->guid.

  endmethod.                    "GET_GUID


  method get_header_footer_drawings.
    data: ls_odd_header  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_odd_footer  type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_even_header type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_even_footer type zif_excel_data_decl=>zexcel_s_worksheet_head_foot,
          ls_hd_ft       type zif_excel_data_decl=>zexcel_s_worksheet_head_foot.

    field-symbols: <fs_drawings> type zif_excel_data_decl=>zexcel_s_drawings.

    me->sheet_setup->get_header_footer( importing ep_odd_header = ls_odd_header
                                                  ep_odd_footer = ls_odd_footer
                                                  ep_even_header = ls_even_header
                                                  ep_even_footer = ls_even_footer ).

**********************************************************************
*** Odd header
    ls_hd_ft = ls_odd_header.
    if ls_hd_ft-left_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    endif.
    if ls_hd_ft-right_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    endif.
    if ls_hd_ft-center_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    endif.

**********************************************************************
*** Odd footer
    ls_hd_ft = ls_odd_footer.
    if ls_hd_ft-left_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    endif.
    if ls_hd_ft-right_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    endif.
    if ls_hd_ft-center_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    endif.

**********************************************************************
*** Even header
    ls_hd_ft = ls_even_header.
    if ls_hd_ft-left_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    endif.
    if ls_hd_ft-right_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    endif.
    if ls_hd_ft-center_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    endif.

**********************************************************************
*** Even footer
    ls_hd_ft = ls_even_footer.
    if ls_hd_ft-left_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    endif.
    if ls_hd_ft-right_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    endif.
    if ls_hd_ft-center_image is not initial.
      append initial line to rt_drawings assigning <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    endif.

  endmethod.                    "get_header_footer_drawings


  method get_highest_column.
    me->update_dimension_range( ).
    r_highest_column = me->lower_cell-cell_column.
  endmethod.                    "GET_HIGHEST_COLUMN


  method get_highest_row.
    me->update_dimension_range( ).
    r_highest_row = me->lower_cell-cell_row.
  endmethod.                    "GET_HIGHEST_ROW


  method get_hyperlinks_iterator.
    eo_iterator = hyperlinks->get_iterator( ).
  endmethod.                    "GET_HYPERLINKS_ITERATOR


  method get_hyperlinks_size.
    ep_size = hyperlinks->size( ).
  endmethod.                    "GET_HYPERLINKS_SIZE


  method get_ignored_errors.
    rt_ignored_errors = mt_ignored_errors.
  endmethod.


  method get_merge.

    field-symbols: <ls_merged_cell> like line of me->mt_merged_cells.

    data: lv_col_from    type string,
          lv_col_to      type string,
          lv_row_from    type string,
          lv_row_to      type string,
          lv_merge_range type string.

    loop at me->mt_merged_cells assigning <ls_merged_cell>.

      lv_col_from = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_from ).
      lv_col_to   = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_to   ).
      lv_row_from = <ls_merged_cell>-row_from.
      lv_row_to   = <ls_merged_cell>-row_to  .
      concatenate lv_col_from lv_row_from ':' lv_col_to lv_row_to
         into lv_merge_range.
      condense lv_merge_range no-gaps.
      append lv_merge_range to merge_range.

    endloop.

  endmethod.                    "GET_MERGE


  method get_pagebreaks.
    ro_pagebreaks = mo_pagebreaks.
  endmethod.                    "GET_PAGEBREAKS


  method get_ranges_iterator.

    eo_iterator = me->ranges->get_iterator( ).

  endmethod.                    "GET_RANGES_ITERATOR


  method get_row.
    eo_row = me->rows->get( ip_index = ip_row ).

    if eo_row is not bound.
      eo_row = me->add_new_row( ip_row ).
    endif.
  endmethod.                    "GET_ROW


  method get_rows.

    data: row type i.
    field-symbols: <sheet_cell> type zif_excel_data_decl=>zexcel_s_cell_data.

    if sheet_content is not initial.

      row = 0.
      do.
        " Find the next row
        read table sheet_content assigning <sheet_cell> with key cell_row = row.
        case sy-subrc.
          when 4.
            " row doesn't exist, but it exists another row, SY-TABIX points to the first cell in this row.
            read table sheet_content assigning <sheet_cell> index sy-tabix.
            assert sy-subrc = 0.
            row = <sheet_cell>-cell_row.
          when 8.
            " it was the last available row
            exit.
        endcase.
        " This will create the row instance if it doesn't exist
        get_row( row ).
        row = row + 1.
      enddo.

    endif.

    eo_rows = me->rows.
  endmethod.                    "GET_ROWS


  method get_rows_iterator.

    get_rows( ).
    eo_iterator = me->rows->get_iterator( ).

  endmethod.                    "GET_ROWS_ITERATOR


  method get_row_outlines.

    rt_row_outlines = me->mt_row_outlines.

  endmethod.                    "GET_ROW_OUTLINES


  method get_style_cond.

    data: lo_style_iterator type ref to zcl_excel_collection_iterator,
          lo_style_cond     type ref to zcl_excel_style_cond.

    lo_style_iterator = me->get_style_cond_iterator( ).
    while lo_style_iterator->has_next( ) = abap_true.
      lo_style_cond ?= lo_style_iterator->get_next( ).
      if lo_style_cond->get_guid( ) = ip_guid.
        eo_style_cond = lo_style_cond.
        exit.
      endif.
    endwhile.

  endmethod.                    "GET_STYLE_COND


  method get_style_cond_iterator.

    eo_iterator = styles_cond->get_iterator( ).
  endmethod.                    "GET_STYLE_COND_ITERATOR


  method get_tabcolor.
    ev_tabcolor = me->tabcolor.
  endmethod.                    "GET_TABCOLOR


  method get_table.
*--------------------------------------------------------------------*
* Comment D. Rauchenstein
* With this method, we get a fully functional Excel Upload, which solves
* a few issues of the other excel upload tools
* ZBCABA_ALSM_EXCEL_UPLOAD_EXT: Reads only up to 50 signs per Cell, Limit
* in row-Numbers. Other have Limitations of Lines, or you are not able
* to ignore filters or choosing the right tab.
*
* To get a fully functional XLSX Upload, you can use it e.g. with method
* CL_EXCEL_READER_2007->ZIF_EXCEL_READER~LOAD_FILE()
*--------------------------------------------------------------------*

    field-symbols: <ls_line> type data.
    field-symbols: <lv_value> type data.

    data lv_actual_row type int4.
    data lv_actual_row_string type string.
    data lv_actual_col type int4.
    data lv_actual_col_string type string.
    data lv_errormessage type string.
    data lv_max_col type zif_excel_data_decl=>zexcel_cell_column.
    data lv_max_row type int4.
    data lv_delta_col type int4.
    data lv_value  type zif_excel_data_decl=>zexcel_cell_value.
    data lv_rc  type sysubrc.
    data lx_conversion_error type ref to cx_sy_conversion_error.
    data lv_float type f.
    data lv_type.
    data lv_tabix type i.

    lv_max_col =  me->get_highest_column( ).
    if iv_max_col is supplied and iv_max_col < lv_max_col.
      lv_max_col = iv_max_col.
    endif.
    lv_max_row =  me->get_highest_row( ).
    if iv_max_row is supplied and iv_max_row < lv_max_row.
      lv_max_row = iv_max_row.
    endif.

*--------------------------------------------------------------------*
* The row counter begins with 1 and should be corrected with the skips
*--------------------------------------------------------------------*
    lv_actual_row =  iv_skipped_rows + 1.
    lv_actual_col =  iv_skipped_cols + 1.


    try.
*--------------------------------------------------------------------*
* Check if we the basic features are possible with given "any table"
*--------------------------------------------------------------------*
        append initial line to et_table assigning <ls_line>.
        if sy-subrc <> 0 or <ls_line> is not assigned.

          lv_errormessage = 'Error at inserting new Line to internal Table'(002).
          zcx_excel=>raise_text( lv_errormessage ).

        else.
          lv_delta_col = lv_max_col - iv_skipped_cols.
          assign component lv_delta_col of structure <ls_line> to <lv_value>.
          if sy-subrc <> 0 or <lv_value> is not assigned.
            lv_errormessage = 'Internal table has less columns than excel'(003).
            zcx_excel=>raise_text( lv_errormessage ).
          else.
*--------------------------------------------------------------------*
*now we are ready for handle the table data
*--------------------------------------------------------------------*
            clear et_table.
*--------------------------------------------------------------------*
* Handle each Row until end on right side
*--------------------------------------------------------------------*
            while lv_actual_row <= lv_max_row .

*--------------------------------------------------------------------*
* Handle each Column until end on bottom
* First step is to step back on first column
*--------------------------------------------------------------------*
              lv_actual_col =  iv_skipped_cols + 1.

              unassign <ls_line>.
              append initial line to et_table assigning <ls_line>.
              if sy-subrc <> 0 or <ls_line> is not assigned.
                lv_errormessage = 'Error at inserting new Line to internal Table'(002).
                zcx_excel=>raise_text( lv_errormessage ).
              endif.
              while lv_actual_col <= lv_max_col.

                lv_delta_col = lv_actual_col - iv_skipped_cols.
                assign component lv_delta_col of structure <ls_line> to <lv_value>.
                if sy-subrc <> 0.
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  concatenate 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string into lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                endif.

                me->get_cell(
                  exporting
                    ip_column  = lv_actual_col    " Cell Column
                    ip_row     = lv_actual_row    " Cell Row
                  importing
                    ep_value   = lv_value    " Cell Value
                    ep_rc      = lv_rc    " Return Value of ABAP Statements
                ).
                if lv_rc <> 0
                  and lv_rc <> 4                                                   "No found error means, zero/no value in cell
                  and lv_rc <> 8. "rc is 8 when the last row contains cells with zero / no values
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  concatenate 'Error at reading field value (Col:'(007) lv_actual_col_string ' Row:'(005) lv_actual_row_string into lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                endif.

                try.
                    data lo_typed type ref to cl_abap_datadescr.
                    lo_typed ?= cl_abap_typedescr=>describe_by_data( <lv_value> ).
                    lv_type = lo_typed->kind.
                    if lv_type = 'D'.
                      <lv_value> = zcl_excel_common=>excel_string_to_date( ip_value = lv_value ).
                    else.
                      <lv_value> = lv_value. "Will raise exception if data type of <lv_value> is not float (or decfloat16/34) and excel delivers exponential number e.g. -2.9398924194538267E-2
                    endif.
                  catch cx_sy_conversion_error into lx_conversion_error.
                    "Another try with conversion to float...
                    if lv_type = 'P'.
                      <lv_value> = lv_float = lv_value.
                    else.
                      raise exception lx_conversion_error. "Pass on original exception
                    endif.
                endtry.

*  CATCH zcx_excel.    "
                lv_actual_col += 1.
              endwhile.
              lv_actual_row += 1.
            endwhile.

            if iv_skip_bottom_empty_rows = abap_true.
              lv_tabix = lines( et_table ).
              while lv_tabix >= 1.
                read table et_table index lv_tabix assigning <ls_line>.
                assert sy-subrc = 0.
                if <ls_line> is not initial.
                  exit.
                endif.
                delete et_table index lv_tabix.
                lv_tabix = lv_tabix - 1.
              endwhile.
            endif.

          endif.


        endif.

      catch cx_sy_assign_cast_illegal_cast.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        concatenate 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string into lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      catch cx_sy_assign_cast_unknown_type.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        concatenate 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string into lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      catch cx_sy_assign_out_of_range.
        lv_errormessage = 'Internal table has less columns than excel'(003).
        zcx_excel=>raise_text( lv_errormessage ).
      catch cx_sy_conversion_error.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        concatenate 'Error at converting field value (Col:'(006) lv_actual_col_string ' Row:'(005) lv_actual_row_string into lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).

    endtry.
  endmethod.                    "get_table


  method get_tables_iterator.
    eo_iterator = tables->get_iterator( ).
  endmethod.                    "GET_TABLES_ITERATOR


  method get_tables_size.
    ep_size = tables->size( ).
  endmethod.                    "GET_TABLES_SIZE


  method get_title.
    data lv_value type string.
    if ip_escaped eq abap_true.
      lv_value = me->title.
      ep_title = zcl_excel_common=>escape_string( lv_value ).
    else.
      ep_title = me->title.
    endif.
  endmethod.                    "GET_TITLE


  method get_value_type.
    data: lo_addit    type ref to cl_abap_elemdescr.

    ep_value = ip_value.
    ep_value_type = cl_abap_typedescr=>typekind_string. " Thats our default if something goes wrong.

    try.
        lo_addit            ?= cl_abap_typedescr=>describe_by_data( ip_value ).
        ep_value_type = lo_addit->type_kind.
      catch cx_sy_move_cast_error.
        clear lo_addit.
    endtry.
  endmethod.                    "GET_VALUE_TYPE


  method is_cell_merged.

    data: lv_column type i.

    field-symbols: <ls_merged_cell> like line of me->mt_merged_cells.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    rp_is_merged = abap_false.                                        " Assume not in merged area

    loop at me->mt_merged_cells assigning <ls_merged_cell>.

      if    <ls_merged_cell>-col_from <= lv_column
        and <ls_merged_cell>-col_to   >= lv_column
        and <ls_merged_cell>-row_from <= ip_row
        and <ls_merged_cell>-row_to   >= ip_row.
        rp_is_merged = abap_true.                                     " until we are proven different
        return.
      endif.

    endloop.

  endmethod.                    "IS_CELL_MERGED


  method move_supplied_borders.

    data: ls_borderx type zif_excel_data_decl=>zexcel_s_cstylex_border.

    if iv_border_supplied = abap_true.  " only act if parameter was supplied
      if iv_xborder_supplied = abap_true. "
        ls_borderx = is_xborder.             " use supplied x-parameter
      else.
        clear ls_borderx with 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        if is_border-border_style is  initial.
          clear ls_borderx-border_style.
        endif.
        clear_initial_colorxfields(
          exporting
            is_color  = is_border-border_color
          changing
            cs_xcolor = ls_borderx-border_color ).
      endif.
      move-corresponding is_border  to cs_complete_style_border.
      move-corresponding ls_borderx to cs_complete_stylex_border.
    endif.

  endmethod.


  method normalize_columnrow_parameter.

    if ( ( ip_column is not initial or ip_row is not initial ) and ip_columnrow is not initial )
        or ( ip_column is initial and ip_row is initial and ip_columnrow is initial ).
      raise exception type zcx_excel
        exporting
          error = 'Please provide either row and column, or cell reference'.
    endif.

    if ip_columnrow is not initial.
      zcl_excel_common=>convert_columnrow2column_a_row(
        exporting
          i_columnrow  = ip_columnrow
        importing
          e_column_int = ep_column
          e_row        = ep_row ).
    else.
      ep_column = zcl_excel_common=>convert_column2int( ip_column ).
      ep_row    = ip_row.
    endif.

  endmethod.


  method normalize_column_heading_texts.
    data: lt_field_catalog      type zif_excel_data_decl=>zexcel_t_fieldcatalog,
          lv_value_lowercase    type string,
          lv_scrtext_l_initial  type zif_excel_data_decl=>zexcel_column_name,
          lv_long_text          type string,
          lv_max_length         type i,
          lv_temp_length        type i,
          lv_syindex            type c length 3,
          lt_column_name_buffer type sorted table of string with unique key table_line.
    field-symbols: <ls_field_catalog> type zif_excel_data_decl=>zexcel_s_fieldcatalog,
                   <scrtxt3>          type any.

    " Due to restrictions in new table object we cannot have two columns with the same name
    " Check if a column with the same name exists, if exists add a counter
    " If no medium description is provided we try to use small or long

    lt_field_catalog = it_field_catalog.

    loop at lt_field_catalog assigning <ls_field_catalog>.

      if <ls_field_catalog>-column_name is initial.

        case iv_default_descr.
          when 'M'.
            assign <ls_field_catalog>-scrtext_l to <scrtxt3>.
          when 'S'.
            assign <ls_field_catalog>-scrtext_l to <scrtxt3>.
          when 'L'.
            assign <ls_field_catalog>-scrtext_l to <scrtxt3>.
          when others.
            assign <ls_field_catalog>-scrtext_l to <scrtxt3>.
        endcase.

        if <scrtxt3> is not initial.
          <ls_field_catalog>-column_name = <scrtxt3>.
        else.
          <ls_field_catalog>-column_name = 'Column'.  " default value as Excel does
        endif.
      endif.

      lv_scrtext_l_initial = <ls_field_catalog>-column_name.
      lv_max_length = strlen( conv string( <ls_field_catalog>-column_name ) ).
      do.
        lv_value_lowercase = <ls_field_catalog>-column_name.
        translate lv_value_lowercase to lower case.
        read table lt_column_name_buffer transporting no fields with key table_line = lv_value_lowercase binary search.
        if sy-subrc <> 0.
          insert lv_value_lowercase into table lt_column_name_buffer.
          exit.
        else.
          lv_syindex = sy-index.
          concatenate lv_scrtext_l_initial lv_syindex into lv_long_text.
          if strlen( lv_long_text ) <= lv_max_length.
            <ls_field_catalog>-column_name = lv_long_text.
          else.
            lv_temp_length = strlen( lv_scrtext_l_initial ) - 1.
            lv_scrtext_l_initial = substring( val = lv_scrtext_l_initial len = lv_temp_length ).
            concatenate lv_scrtext_l_initial lv_syindex into <ls_field_catalog>-column_name.
          endif.
        endif.
      enddo.

    endloop.

    result = lt_field_catalog.

  endmethod.


  method normalize_range_parameter.

    data: lv_errormessage type string.

    if ( ( ip_column_start is not initial or ip_column_end is not initial
            or ip_row is not initial or ip_row_to is not initial ) and ip_range is not initial )
        or ( ip_column_start is initial and ip_column_end is initial
            and ip_row is initial and ip_row_to is initial and ip_range is initial ).
      raise exception type zcx_excel
        exporting
          error = 'Please provide either row and column interval, or range reference'.
    endif.

    if ip_range is not initial.
      zcl_excel_common=>convert_range2column_a_row(
        exporting
          i_range            = ip_range
        importing
          e_column_start_int = ep_column_start
          e_column_end_int   = ep_column_end
          e_row_start        = ep_row
          e_row_end          = ep_row_to ).
    else.
      if ip_column_start is initial.
        ep_column_start = zcl_excel_common=>c_excel_sheet_min_col.
      else.
        ep_column_start = zcl_excel_common=>convert_column2int( ip_column_start ).
      endif.
      if ip_column_end is initial.
        ep_column_end = ep_column_start.
      else.
        ep_column_end = zcl_excel_common=>convert_column2int( ip_column_end ).
      endif.
      ep_row = ip_row.
      if ep_row is initial.
        ep_row = zcl_excel_common=>c_excel_sheet_min_row.
      endif.
      ep_row_to = ip_row_to.
      if ep_row_to is initial.
        ep_row_to = ep_row.
      endif.
    endif.

    if ep_row > ep_row_to.
      lv_errormessage = 'First row larger than last row'(405).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if ep_column_start > ep_column_end.
      lv_errormessage = 'First column larger than last column'(406).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

  endmethod.


  method normalize_style_parameter.

    data lo_style_type type ref to cl_abap_typedescr.
    field-symbols <style> type ref to zcl_excel_style.

    check ip_style_or_guid is not initial.

    lo_style_type = cl_abap_typedescr=>describe_by_data( ip_style_or_guid ).
    if lo_style_type->type_kind = lo_style_type->typekind_oref.
      assign ip_style_or_guid to <style>.
      rv_guid = <style>->get_guid( ).
    elseif lo_style_type->type_kind = lo_style_type->typekind_hex or lo_style_type->type_kind = lo_style_type->typekind_xstring.
      rv_guid = ip_style_or_guid.
    else.
      raise exception type zcx_excel exporting error = 'IP_GUID type must be either REF TO zcl_excel_style or a HEX value'.
    endif.

  endmethod.


  method print_title_set_range.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


    data: lo_range_iterator         type ref to zcl_excel_collection_iterator,
          lo_range                  type ref to zcl_excel_range,
          lv_repeat_range_sheetname type string,
          lv_repeat_range_col       type string,
          lv_row_char_from          type c length 10,
          lv_row_char_to            type c length 10,
          lv_repeat_range_row       type string,
          lv_repeat_range           type string.


*--------------------------------------------------------------------*
* Get range that represents printarea
* if non-existant, create it
*--------------------------------------------------------------------*
    lo_range_iterator = me->get_ranges_iterator( ).
    while lo_range_iterator->has_next( ) = abap_true.

      lo_range ?= lo_range_iterator->get_next( ).
      if lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
        exit.  " Found it
      endif.
      clear lo_range.

    endwhile.


    if me->print_title_col_from is initial and
       me->print_title_row_from is initial.
*--------------------------------------------------------------------*
* No print titles are present,
*--------------------------------------------------------------------*
      if lo_range is bound.
        me->ranges->remove( lo_range ).
      endif.
    else.
*--------------------------------------------------------------------*
* Print titles are present,
*--------------------------------------------------------------------*
      if lo_range is not bound.
        lo_range =  me->add_new_range( ).
        lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
      endif.

      lv_repeat_range_sheetname = me->get_title( ).
      lv_repeat_range_sheetname = zcl_excel_common=>escape_string( lv_repeat_range_sheetname ).

*--------------------------------------------------------------------*
* Repeat-columns
*--------------------------------------------------------------------*
      if me->print_title_col_from is not initial.
        concatenate lv_repeat_range_sheetname
                    '!$' me->print_title_col_from
                    ':$' me->print_title_col_to
            into lv_repeat_range_col.
      endif.

*--------------------------------------------------------------------*
* Repeat-rows
*--------------------------------------------------------------------*
      if me->print_title_row_from is not initial.
        lv_row_char_from = me->print_title_row_from.
        lv_row_char_to   = me->print_title_row_to.
        concatenate '!$' lv_row_char_from
                    ':$' lv_row_char_to
            into lv_repeat_range_row.
        condense lv_repeat_range_row no-gaps.
        concatenate lv_repeat_range_sheetname
                    lv_repeat_range_row
            into lv_repeat_range_row.
      endif.

*--------------------------------------------------------------------*
* Concatenate repeat-rows and columns
*--------------------------------------------------------------------*
      if lv_repeat_range_col is initial.
        lv_repeat_range = lv_repeat_range_row.
      elseif lv_repeat_range_row is initial.
        lv_repeat_range = lv_repeat_range_col.
      else.
        concatenate lv_repeat_range_col lv_repeat_range_row
            into lv_repeat_range separated by ','.
      endif.


      lo_range->set_range_value( lv_repeat_range ).
    endif.



  endmethod.                    "PRINT_TITLE_SET_RANGE


  method set_area.

    data: lv_row              type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_start        type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_end          type zif_excel_data_decl=>zexcel_cell_row,
          lv_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          lv_column           type zif_excel_data_decl=>zexcel_cell_column_alpha,
          lv_column_start_int type zif_excel_data_decl=>zexcel_cell_column,
          lv_column_end_int   type zif_excel_data_decl=>zexcel_cell_column.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               importing ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    if ip_merge = abap_true or ip_area = c_area-topleft.

      if ip_data_type is supplied or
         ip_abap_type is supplied.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink
                      ip_data_type = ip_data_type
                      ip_abap_type = ip_abap_type ).

      else.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink ).

      endif.

    else.

      lv_column_int = lv_column_start_int.
      while lv_column_int <= lv_column_end_int.

        lv_column = zcl_excel_common=>convert_column2alpha( lv_column_int ).
        lv_row = lv_row_start.

        while lv_row <= lv_row_end.

          if ip_data_type is supplied or
             ip_abap_type is supplied.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink
                          ip_data_type = ip_data_type
                          ip_abap_type = ip_abap_type ).

          else.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink ).

          endif.

          lv_row += 1.
        endwhile.

        lv_column_int += 1.
      endwhile.

    endif.

    if ip_style is supplied.

      me->set_area_style( ip_column_start = lv_column_start_int
                          ip_column_end   = lv_column_end_int
                          ip_row          = lv_row_start
                          ip_row_to       = lv_row_end
                          ip_style        = ip_style ).
    endif.

    if ip_merge is supplied and ip_merge = abap_true.

      me->set_merge( ip_column_start = lv_column_start_int
                     ip_column_end   = lv_column_end_int
                     ip_row          = lv_row_start
                     ip_row_to       = lv_row_end ).

    endif.

  endmethod.                    "set_area


  method set_area_formula.
    data: ld_row              type zif_excel_data_decl=>zexcel_cell_row,
          ld_row_start        type zif_excel_data_decl=>zexcel_cell_row,
          ld_row_end          type zif_excel_data_decl=>zexcel_cell_row,
          ld_column           type zif_excel_data_decl=>zexcel_cell_column_alpha,
          ld_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_start_int type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_end_int   type zif_excel_data_decl=>zexcel_cell_column.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               importing ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    if ip_merge = abap_true or ip_area = c_area-topleft.

      me->set_cell_formula( ip_column = ld_column_start_int ip_row = ld_row_start
                            ip_formula = ip_formula ).

    else.

      ld_column_int = ld_column_start_int.
      while ld_column_int <= ld_column_end_int.

        ld_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
        ld_row = ld_row_start.
        while ld_row <= ld_row_end.

          me->set_cell_formula( ip_column = ld_column ip_row = ld_row
                                ip_formula = ip_formula ).

          ld_row += 1.
        endwhile.

        ld_column_int += 1.
      endwhile.

    endif.

    if ip_merge is supplied and ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    endif.
  endmethod.                    "set_area_formula


  method set_area_hyperlink.
    data: ld_row_start        type zif_excel_data_decl=>zexcel_cell_row,
          ld_row_end          type zif_excel_data_decl=>zexcel_cell_row,
          ld_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_start_int type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_end_int   type zif_excel_data_decl=>zexcel_cell_column,
          ld_current_column   type zif_excel_data_decl=>zexcel_cell_column_alpha,
          ld_current_row      type zif_excel_data_decl=>zexcel_cell_row,
          ld_value            type string,
          ld_formula          type string.
    data: lo_hyperlink type ref to zcl_excel_hyperlink.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               importing ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    ld_column_int = ld_column_start_int.
    while ld_column_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      while ld_current_row <= ld_row_end.

        me->get_cell( exporting ip_column  = ld_current_column ip_row = ld_current_row
                      importing ep_value   = ld_value
                                ep_formula = ld_formula ).

        if ip_is_internal = abap_true.
          lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = ip_url ).
        else.
          lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = ip_url ).
        endif.

        me->set_cell( ip_column = ld_current_column ip_row = ld_current_row ip_value = ld_value ip_formula = ld_formula ip_hyperlink = lo_hyperlink ).

        ld_current_row += 1.
      endwhile.
      ld_column_int += 1.
    endwhile.

  endmethod.                    "SET_AREA_HYPERLINK


  method set_area_style.
    data: ld_row_start        type zif_excel_data_decl=>zexcel_cell_row,
          ld_row_end          type zif_excel_data_decl=>zexcel_cell_row,
          ld_column_int       type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_start_int type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_end_int   type zif_excel_data_decl=>zexcel_cell_column,
          ld_current_column   type zif_excel_data_decl=>zexcel_cell_column_alpha,
          ld_current_row      type zif_excel_data_decl=>zexcel_cell_row.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               importing ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    ld_column_int = ld_column_start_int.
    while ld_column_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      while ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ld_current_row += 1.
      endwhile.
      ld_column_int += 1.
    endwhile.
    if ip_merge is supplied and ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    endif.
  endmethod.                    "SET_AREA_STYLE


  method set_cell.
    data lv_column        type zif_excel_data_decl=>zexcel_cell_column.
    data ls_sheet_content type zif_excel_data_decl=>zexcel_s_cell_data.
    data lv_row           type zif_excel_data_decl=>zexcel_cell_row.
    data lv_value         type zif_excel_data_decl=>zexcel_cell_value.
    data lv_data_type     type zif_excel_data_decl=>zexcel_cell_data_type.
    data lv_value_type    type abap_typekind.
    data lv_style_guid    type zif_excel_data_decl=>zexcel_cell_style.
    data lo_addit         type ref to cl_abap_elemdescr.
    data lo_type          type ref to cl_abap_datadescr.
    data lt_rtf           type zif_excel_data_decl=>zexcel_t_rtf.
    data lo_value         type ref to data.
    data lo_value_new     type ref to data.
    data lv_newformat type zif_excel_data_decl=>zexcel_number_format.
    field-symbols <fs_sheet_content>  type zif_excel_data_decl=>zexcel_s_cell_data.
    field-symbols <fs_numeric>        type numeric.
    field-symbols <fs_date>           type d.
    field-symbols <fs_time>           type t.
    field-symbols <fs_value>          type simple.
    field-symbols <fs_typekind_int8>  type abap_typekind.
    field-symbols <fs_column_formula> type mty_s_column_formula.
    field-symbols <ls_fieldcat>       type zif_excel_data_decl=>zexcel_s_fieldcatalog.
    field-symbols <lv_utclong>        type simple.

    if     ip_value             is not supplied
       and ip_formula           is not supplied
       and ip_column_formula_id  = 0.
      zcx_excel=>raise_text( 'Please provide the value or formula' ).
    endif.

    normalize_columnrow_parameter( exporting ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   importing ep_column    = lv_column
                                             ep_row       = lv_row ).

    " Begin of change issue #152 - don't touch existing style if only value is passed
    if ip_column_formula_id <> 0.
      check_cell_column_formula( it_column_formulas   = column_formulas
                                 ip_column_formula_id = ip_column_formula_id
                                 ip_formula           = ip_formula
                                 ip_value             = ip_value
                                 ip_row               = lv_row
                                 ip_column            = lv_column ).
    endif.
    assign sheet_content[ cell_row    = lv_row      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                          cell_column = lv_column ] to <fs_sheet_content>.
    if sy-subrc = 0.
      if ip_style is initial.
        " If no style is provided as method-parameter and cell is found use cell's current style
        lv_style_guid = <fs_sheet_content>-cell_style.
      else.
        " Style provided as method-parameter --> use this
        lv_style_guid = normalize_style_parameter( ip_style ).
      endif.
    else.
      " No cell found --> use supplied style even if empty
      lv_style_guid = normalize_style_parameter( ip_style ).
    endif.
    " End of change issue #152 - don't touch existing style if only value is passed

    if ip_value is not supplied.
      return.
    endif.

    " if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
    " IP_DATA_TYPE is passed by excel reader so source types are preserved
    " First we get reference into local var.
    lo_type ?= cl_abap_datadescr=>describe_by_data( ip_value ).
    try.
        create data lo_value type handle lo_type.
      catch cx_sy_create_data_error.
        create data lo_value type string.
    endtry.

    assign lo_value->* to <fs_value>.
    if sy-subrc = 0.
      <fs_value> = ip_value.
      if ip_data_type is supplied.
        if ip_abap_type is not supplied.
          get_value_type( exporting ip_value = ip_value
                          importing ep_value = <fs_value> ).
        endif.
        lv_value = <fs_value>.
        lv_data_type = ip_data_type.
      else.
        if ip_abap_type is supplied.
          lv_value_type = ip_abap_type.
        else.
          get_value_type( exporting ip_value      = ip_value
                          importing ep_value      = <fs_value>
                                    ep_value_type = lv_value_type ).
        endif.

        assign cl_abap_typedescr=>typekind_int8 to <fs_typekind_int8>.

        case lv_value_type.
          when cl_abap_typedescr=>typekind_int or cl_abap_typedescr=>typekind_int1 or cl_abap_typedescr=>typekind_int2
            or <fs_typekind_int8>. " Allow INT8 types columns
            if lv_value_type = <fs_typekind_int8>.
              call method cl_abap_elemdescr=>('GET_INT8')
                receiving
                  p_result = lo_addit.
            else.
              lo_addit = cl_abap_elemdescr=>get_i( ).
            endif.
            create data lo_value_new type handle lo_addit.
            assign lo_value_new->* to <fs_numeric>.
            if sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            endif.

          when cl_abap_typedescr=>typekind_float or cl_abap_typedescr=>typekind_packed or
               cl_abap_typedescr=>typekind_decfloat or
               cl_abap_typedescr=>typekind_decfloat16 or
               cl_abap_typedescr=>typekind_decfloat34.
            if     lv_value_type  = cl_abap_typedescr=>typekind_packed
               and ip_currency   is not initial.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value    = <fs_value>
                                                                   ip_currency = ip_currency ).
            elseif     lv_value_type     = cl_abap_typedescr=>typekind_packed
                   and ip_unitofmeasure is not initial.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value         = <fs_value>
                                                                   ip_unitofmeasure = ip_unitofmeasure ).
            else.
              lo_addit = cl_abap_elemdescr=>get_f( ).
              create data lo_value_new type handle lo_addit.
              assign lo_value_new->* to <fs_numeric>.
              if sy-subrc = 0.
                <fs_numeric> = <fs_value>.
                lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              endif.
            endif.

          when cl_abap_typedescr=>typekind_char or cl_abap_typedescr=>typekind_string or cl_abap_typedescr=>typekind_num or
               cl_abap_typedescr=>typekind_hex or cl_abap_typedescr=>typekind_xstring.
            lv_value = <fs_value>.
            lv_data_type = 's'.

          when cl_abap_typedescr=>typekind_date.
            lo_addit = cl_abap_elemdescr=>get_d( ).
            create data lo_value_new type handle lo_addit.
            assign lo_value_new->* to <fs_date>.
            if sy-subrc = 0.
              <fs_date> = <fs_value>.
              lv_value = zcl_excel_common=>date_to_excel_string( ip_value = <fs_date> ).
            endif.
* Begin of change issue #152 - don't touch existing style if only value is passed
* Moved to end of routine - apply date-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default date format in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = get_default_excel_date_format( ).
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch existing style if only value is passed

          when cl_abap_typedescr=>typekind_time.
            lo_addit = cl_abap_elemdescr=>get_t( ).
            create data lo_value_new type handle lo_addit.
            assign lo_value_new->* to <fs_time>.
            if sy-subrc = 0.
              <fs_time> = <fs_value>.
              lv_value = zcl_excel_common=>time_to_excel_string( ip_value = <fs_time> ).
            endif.
* Begin of change issue #152 - don't touch existing style if only value is passed
* Moved to end of routine - apply time-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default time format for user in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_time6.
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch existing style if only value is passed

          when typekind_utclong.
            assign variable_utclong->* to <lv_utclong>.
            if sy-subrc = 0.
              <lv_utclong> = <fs_value>.
              lv_value = zcl_excel_common=>utclong_to_excel_string( <lv_utclong> ).
            endif.

          when others.
            zcx_excel=>raise_text( 'Invalid data type of input value' ).
        endcase.
      endif.

      if <fs_sheet_content> is assigned and <fs_sheet_content>-table_header is not initial and lv_value is not initial.
        assign <fs_sheet_content>-table->fieldcat[ fieldname = <fs_sheet_content>-table_fieldname ] to <ls_fieldcat>.
        if sy-subrc = 0.
          <ls_fieldcat>-column_name = lv_value.
          if <ls_fieldcat>-column_name <> lv_value.
            zcx_excel=>raise_text( 'Cell is table column header - this value is not allowed' ).
          endif.
        endif.
      endif.

    endif.

    if ip_hyperlink is bound.
      ip_hyperlink->set_cell_reference( ip_column = lv_column
                                        ip_row    = lv_row ).
      hyperlinks->add( ip_hyperlink ).
    endif.

    if lv_value cs '_x'.
      " Issue #761 value "_x0041_" rendered as "A".
      " "_x...._", where "." is 0-9 a-f or A-F (case insensitive), is an internal value in sharedStrings.xml
      " that Excel uses to store special characters, it's interpreted like Unicode character U+....
      " for instance "_x0041_" is U+0041 which is "A".
      " To not interpret such text, the first underscore is replaced with "_x005f_".
      " The value "_x0041_" is to be stored internally "_x005f_x0041_" so that it's rendered like "_x0041_".
      " Note that REGEX is time consuming, it's why "CS" is used above to improve the performance.
      replace all occurrences of pcre '_(x[0-9a-fA-F]{4}_)' in lv_value with '_x005f_$1' respecting case.
    endif.

    " Begin of change issue #152 - don't touch existing style if only value is passed
    " Read table moved up, so that current style may be evaluated

    if ip_textvalue is not initial.
      lv_value = |{ lv_value } ({ ip_textvalue })|.
    endif.

    if <fs_sheet_content> is assigned.
      " End of change issue #152 - don't touch existing style if only value is passed
      <fs_sheet_content>-cell_value        = lv_value.
      <fs_sheet_content>-cell_formula      = ip_formula.
      <fs_sheet_content>-column_formula_id = ip_column_formula_id.
      <fs_sheet_content>-cell_style        = lv_style_guid.
      <fs_sheet_content>-data_type         = lv_data_type.
    else.
      ls_sheet_content-cell_row          = lv_row.
      ls_sheet_content-cell_column       = lv_column.
      ls_sheet_content-cell_value        = lv_value.
      ls_sheet_content-cell_formula      = ip_formula.
      ls_sheet_content-column_formula_id = ip_column_formula_id.
      ls_sheet_content-cell_style        = lv_style_guid.
      ls_sheet_content-data_type         = lv_data_type.
      ls_sheet_content-cell_coords       = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column
                                                                                             i_row    = lv_row ).
      insert ls_sheet_content into table sheet_content assigning <fs_sheet_content>. " ins #152 - Now <fs_sheet_content> always holds the data

    endif.

    if ip_formula is initial and lv_value is not initial and it_rtf is not initial.
      lt_rtf = it_rtf.
      check_rtf( exporting ip_value = lv_value
                           ip_style = lv_style_guid
                 changing  ct_rtf   = lt_rtf ).
      <fs_sheet_content>-rtf_tab = lt_rtf.
    endif.

    " Begin of change issue #152 - don't touch existing style if only value is passed
    " For Date- or Timefields change the formatcode if nothing is set yet
    " Enhancement option:  Check if existing formatcode is a date/ or timeformat
    "                      If not, use default
    data lo_format_code_datetime type zif_excel_data_decl=>zexcel_number_format.
    data stylemapping            type zif_excel_data_decl=>zexcel_s_stylemapping.
    if <fs_sheet_content>-cell_style is initial.
      <fs_sheet_content>-cell_style = excel->get_default_style( ).
    endif.

    case lv_value_type.
      when cl_abap_typedescr=>typekind_date.
        try.
            stylemapping = excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          catch zcx_excel.
        endtry.
        if    stylemapping-complete_stylex-number_format-format_code is initial
           or stylemapping-complete_style-number_format-format_code  is initial.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_std.
        else.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        endif.
        change_cell_style( ip_column                    = lv_column
                           ip_row                       = lv_row
                           ip_number_format_format_code = lo_format_code_datetime ).

      when cl_abap_typedescr=>typekind_time.
        try.
            stylemapping = excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          catch zcx_excel.
        endtry.
        if    stylemapping-complete_stylex-number_format-format_code is initial
           or stylemapping-complete_style-number_format-format_code  is initial.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_time6.
        else.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        endif.
        change_cell_style( ip_column                    = lv_column
                           ip_row                       = lv_row
                           ip_number_format_format_code = lo_format_code_datetime ).

      when typekind_utclong.
        try.
            stylemapping = excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          catch zcx_excel.
        endtry.
        if    stylemapping-complete_stylex-number_format-format_code is initial
           or stylemapping-complete_style-number_format-format_code  is initial.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_datetime.
        else.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        endif.
        change_cell_style( ip_column                    = lv_column
                           ip_row                       = lv_row
                           ip_number_format_format_code = lo_format_code_datetime ).

    endcase.
    " End of change issue #152 - don't touch existing style if only value is passed

    " Fix issue #162
    lv_value = ip_value.
    if lv_value cs cl_abap_char_utilities=>cr_lf.
      change_cell_style( ip_column             = lv_column
                         ip_row                = lv_row
                         ip_alignment_wraptext = abap_true ).
    endif.

    if ip_currency is not initial or ip_unitofmeasure is not initial.
      clear lv_newformat.
      if ip_currency is not initial.
        read table zcl_excel_common=>lt_currs into data(ls_curr) with key curr = ip_currency binary search.
        lv_newformat = |* #,##0{ cond string( when ls_curr-dec is not initial then
          '.' && repeat( val = '0' occ = ls_curr-dec )
          else space ) }" { ip_currency }"|.
      elseif ip_unitofmeasure is not initial.
        read table zcl_excel_common=>lt_uoms into data(ls_uom) with key uom = ip_unitofmeasure binary search.
        lv_newformat = |* #,##0{ cond string( when ls_uom-dec is not initial then
          '.' && repeat( val = '0' occ = ls_uom-dec )
          else space ) }" { ls_uom-uome }"|.
      endif.
      if lv_newformat is not initial.
        change_cell_style( ip_column = lv_column
                           ip_row    = lv_row
                           ip_number_format_format_code = lv_newformat ).
      endif.
    endif.
    " End of Fix issue #162
  endmethod.


  method set_cell_formula.
    data:
      lv_column        type zif_excel_data_decl=>zexcel_cell_column,
      lv_row           type zif_excel_data_decl=>zexcel_cell_row,
      ls_sheet_content like line of me->sheet_content.

    field-symbols:
                <sheet_content>                 like line of me->sheet_content.

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
    normalize_columnrow_parameter( exporting ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   importing ep_column    = lv_column
                                             ep_row       = lv_row ).

    read table me->sheet_content assigning <sheet_content> with table key cell_row    = lv_row
                                                                          cell_column = lv_column.
    if sy-subrc <> 0.                   " Create new entry in sheet_content if necessary
      check ip_formula is not initial.  " only create new entry in sheet_content when a formula is passed
      ls_sheet_content-cell_row    = lv_row.
      ls_sheet_content-cell_column = lv_column.
      ls_sheet_content-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column i_row = lv_row ).
      insert ls_sheet_content into table me->sheet_content assigning <sheet_content>.
    endif.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
    <sheet_content>-cell_formula = ip_formula.


  endmethod.                    "SET_CELL_FORMULA


  method set_cell_style.

    data: lv_column     type zif_excel_data_decl=>zexcel_cell_column,
          lv_row        type zif_excel_data_decl=>zexcel_cell_row,
          lv_style_guid type zif_excel_data_decl=>zexcel_cell_style.

    field-symbols: <fs_sheet_content> type zif_excel_data_decl=>zexcel_s_cell_data.

    lv_style_guid = normalize_style_parameter( ip_style ).

    normalize_columnrow_parameter( exporting ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   importing ep_column    = lv_column
                                             ep_row       = lv_row ).

    read table sheet_content assigning <fs_sheet_content> with key cell_row    = lv_row
                                                                   cell_column = lv_column.

    if sy-subrc eq 0.
      <fs_sheet_content>-cell_style   = lv_style_guid.
    else.
      set_cell( ip_column = ip_column ip_row = ip_row ip_value = '' ip_style = ip_style ).
    endif.

  endmethod.                    "SET_CELL_STYLE


  method set_column_width.
    data: lo_column  type ref to zcl_excel_column.
    data: width             type f.

    lo_column = me->get_column( ip_column ).

* if a fix size is supplied use this
    if ip_width_fix is supplied.
      try.
          width = ip_width_fix.
          if width <= 0.
            zcx_excel=>raise_text( 'Please supply a positive number as column-width' ).
          endif.
          lo_column->set_width( width ).
          return.
        catch cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
          zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
      endtry.
    endif.

* If we get down to here, we have to use whatever is found in autosize.
    lo_column->set_auto_size( ip_width_autosize ).


  endmethod.                    "SET_COLUMN_WIDTH


  method set_default_excel_date_format.

    if ip_default_excel_date_format is initial.
      zcx_excel=>raise_text( 'Default date format cannot be blank' ).
    endif.

    default_excel_date_format = ip_default_excel_date_format.
  endmethod.                    "SET_DEFAULT_EXCEL_DATE_FORMAT


  method set_ignored_errors.
    mt_ignored_errors = it_ignored_errors.
  endmethod.


  method set_merge.

    data: ls_merge        type mty_merge,
          lv_column_start type zif_excel_data_decl=>zexcel_cell_column,
          lv_column_end   type zif_excel_data_decl=>zexcel_cell_column,
          lv_row          type zif_excel_data_decl=>zexcel_cell_row,
          lv_row_to       type zif_excel_data_decl=>zexcel_cell_row,
          lv_errormessage type string.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start ip_column_end = ip_column_end
                                         ip_row          = ip_row          ip_row_to     = ip_row_to
                               importing ep_column_start = lv_column_start ep_column_end = lv_column_end
                                         ep_row          = lv_row          ep_row_to     = lv_row_to ).

    if ip_value is supplied or ip_formula is supplied.
      " if there is a value or formula set the value to the top-left cell
      "maybe it is necessary to support other paramters for set_cell
      if ip_value is supplied.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_value ).
      endif.
      if ip_formula is supplied.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_formula ).
      endif.
    endif.
    "call to set_merge_style to apply the style to all cells at the matrix
    if ip_style is supplied.
      me->set_merge_style( ip_row = lv_row ip_column_start = lv_column_start
                           ip_row_to = lv_row_to ip_column_end = lv_column_end
                           ip_style = ip_style ).
    endif.
    ...
*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
    ls_merge-row_from = lv_row.
    ls_merge-row_to   = lv_row_to.
    ls_merge-col_from = lv_column_start.
    ls_merge-col_to   = lv_column_end.

*--------------------------------------------------------------------*
* Check merge not overlapping with existing merges
*--------------------------------------------------------------------*
    loop at me->mt_merged_cells transporting no fields where not (    row_from > ls_merge-row_to
                                                                   or row_to   < ls_merge-row_from
                                                                   or col_from > ls_merge-col_to
                                                                   or col_to   < ls_merge-col_from ). "#EC CI_SORTSEQ
      lv_errormessage = 'Overlapping merges'(404).
      zcx_excel=>raise_text( lv_errormessage ).

    endloop.

*--------------------------------------------------------------------*
* Everything seems ok --> add to merge table
*--------------------------------------------------------------------*
    insert ls_merge into table me->mt_merged_cells.

  endmethod.                    "SET_MERGE


  method set_merge_style.
    data: ld_row_start      type zif_excel_data_decl=>zexcel_cell_row,
          ld_row_end        type zif_excel_data_decl=>zexcel_cell_row,
          ld_column_int     type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_start   type zif_excel_data_decl=>zexcel_cell_column,
          ld_column_end     type zif_excel_data_decl=>zexcel_cell_column,
          ld_current_column type zif_excel_data_decl=>zexcel_cell_column_alpha,
          ld_current_row    type zif_excel_data_decl=>zexcel_cell_row.

    normalize_range_parameter( exporting ip_range        = ip_range
                                         ip_column_start = ip_column_start ip_column_end = ip_column_end
                                         ip_row          = ip_row          ip_row_to     = ip_row_to
                               importing ep_column_start = ld_column_start ep_column_end = ld_column_end
                                         ep_row          = ld_row_start    ep_row_to     = ld_row_end ).

    "set the style cell by cell
    ld_column_int = ld_column_start.
    while ld_column_int <= ld_column_end.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      while ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ld_current_row += 1.
      endwhile.
      ld_column_int += 1.
    endwhile.
  endmethod.                    "set_merge_style


  method set_pane_top_left_cell.
    data lv_column_int type zif_excel_data_decl=>zexcel_cell_column.
    data lv_row type zif_excel_data_decl=>zexcel_cell_row.

    " Validate input value
    zcl_excel_common=>convert_columnrow2column_a_row(
      exporting
        i_columnrow  = iv_columnrow
      importing
        e_column_int = lv_column_int
        e_row        = lv_row ).
    if lv_column_int not between zcl_excel_common=>c_excel_sheet_min_col and zcl_excel_common=>c_excel_sheet_max_col
        or lv_row not between zcl_excel_common=>c_excel_sheet_min_row and zcl_excel_common=>c_excel_sheet_max_row.
      raise exception type zcx_excel exporting error = 'Invalid column/row coordinates (valid values: A1 to XFD1048576)'.
    endif.
    pane_top_left_cell = iv_columnrow.
  endmethod.


  method set_print_gridlines.
    me->print_gridlines = i_print_gridlines.
  endmethod.                    "SET_PRINT_GRIDLINES


  method set_row_height.
    data: lo_row  type ref to zcl_excel_row.
    data: height  type f.

    lo_row = me->get_row( ip_row ).

* if a fix size is supplied use this
    try.
        height = ip_height_fix.
        lo_row->set_row_height( height ).
        return.
      catch cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
        zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
    endtry.

  endmethod.                    "SET_ROW_HEIGHT


  method set_row_outline.

    data: ls_row_outline like line of me->mt_row_outlines.
    field-symbols: <ls_row_outline> like line of me->mt_row_outlines.

    read table me->mt_row_outlines assigning <ls_row_outline> with table key row_from = iv_row_from
                                                                             row_to   = iv_row_to.
    if sy-subrc <> 0.
      if iv_row_from <= 0.
        zcx_excel=>raise_text( 'First row of outline must be a positive number' ).
      endif.
      if iv_row_to < iv_row_from.
        zcx_excel=>raise_text( 'Last row of outline may not be less than first line of outline' ).
      endif.
      ls_row_outline-row_from = iv_row_from.
      ls_row_outline-row_to   = iv_row_to.
      insert ls_row_outline into table me->mt_row_outlines assigning <ls_row_outline>.
    endif.

    case iv_collapsed.

      when abap_true
        or abap_false.
        <ls_row_outline>-collapsed = iv_collapsed.

      when others.
        zcx_excel=>raise_text( 'Unknown collapse state' ).

    endcase.
  endmethod.                    "SET_ROW_OUTLINE


  method set_sheetview_top_left_cell.
    data lv_column_int type zif_excel_data_decl=>zexcel_cell_column.
    data lv_row type zif_excel_data_decl=>zexcel_cell_row.

    " Validate input value
    zcl_excel_common=>convert_columnrow2column_a_row(
      exporting
        i_columnrow  = iv_columnrow
      importing
        e_column_int = lv_column_int
        e_row        = lv_row ).
    if lv_column_int not between zcl_excel_common=>c_excel_sheet_min_col and zcl_excel_common=>c_excel_sheet_max_col
        or lv_row not between zcl_excel_common=>c_excel_sheet_min_row and zcl_excel_common=>c_excel_sheet_max_row.
      raise exception type zcx_excel exporting error = 'Invalid column/row coordinates (valid values: A1 to XFD1048576)'.
    endif.
    sheetview_top_left_cell = iv_columnrow.
  endmethod.


  method set_show_gridlines.
    me->show_gridlines = i_show_gridlines.
  endmethod.                    "SET_SHOW_GRIDLINES


  method set_show_rowcolheaders.
    me->show_rowcolheaders = i_show_rowcolheaders.
  endmethod.                    "SET_SHOW_ROWCOLHEADERS


  method set_tabcolor.
    me->tabcolor = iv_tabcolor.
  endmethod.                    "SET_TABCOLOR


  method set_table.
    "@TODO: to be Fixed by Juwin
*    DATA: lo_structdescr  TYPE REF TO cl_abap_structdescr,
*          lr_data         TYPE REF TO data,
*          lt_dfies        TYPE ddfields,
*          lv_row_int      type zif_excel_data_decl=>zexcel_cell_row,
*          lv_column_int   type zif_excel_data_decl=>zexcel_cell_column,
*          lv_column_alpha type zif_excel_data_decl=>zexcel_cell_column_alpha,
*          lv_cell_value   type zif_excel_data_decl=>zexcel_cell_value.
*
*
*    FIELD-SYMBOLS: <fs_table_line> TYPE any,
*                   <fs_fldval>     TYPE any,
*                   <fs_dfies>      TYPE dfies.
*
*    lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
*    lv_row_int    = ip_top_left_row.
*
*    CREATE DATA lr_data LIKE LINE OF ip_table.
*
*    lo_structdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).
*
*    lt_dfies = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).
*
** It is better to loop column by column
*    LOOP AT lt_dfies ASSIGNING <fs_dfies>.
*      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).
*
*      IF ip_no_header = abap_false.
*        " First of all write column header
*        lv_cell_value = <fs_dfies>-scrtext_m.
*        me->set_cell( ip_column = lv_column_alpha
*                      ip_row    = lv_row_int
*                      ip_value  = lv_cell_value
*                      ip_style  = ip_hdr_style ).
*        IF ip_transpose = abap_true.
*          lv_column_int += 1.
*        ELSE.
*          lv_row_int += 1.
*        ENDIF.
*      ENDIF.
*
*      LOOP AT ip_table ASSIGNING <fs_table_line>.
*        lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).
*        ASSIGN COMPONENT <fs_dfies>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.
*        lv_cell_value = <fs_fldval>.
*        me->set_cell( ip_column = lv_column_alpha
*                      ip_row    = lv_row_int
*                      ip_value  = <fs_fldval>   "lv_cell_value
*                      ip_style  = ip_body_style ).
*        IF ip_transpose = abap_true.
*          lv_column_int += 1.
*        ELSE.
*          lv_row_int += 1.
*        ENDIF.
*      ENDLOOP.
*      IF ip_transpose = abap_true.
*        lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
*        lv_row_int += 1.
*      ELSE.
*        lv_row_int = ip_top_left_row.
*        lv_column_int += 1.
*      ENDIF.
*    ENDLOOP.

  endmethod.                    "SET_TABLE


  method set_table_reference.

    field-symbols: <ls_sheet_content> type zif_excel_data_decl=>zexcel_s_cell_data.

    read table sheet_content assigning <ls_sheet_content> with key cell_row    = ip_row
                                                                   cell_column = ip_column.
    if sy-subrc = 0.
      <ls_sheet_content>-table           = ir_table.
      <ls_sheet_content>-table_fieldname = ip_fieldname.
      <ls_sheet_content>-table_header    = ip_header.
    else.
      zcx_excel=>raise_text( 'Cell not found' ).
    endif.

  endmethod.


  method set_title.
    data: lo_worksheets_iterator type ref to zcl_excel_collection_iterator,
          lo_worksheet           type ref to zcl_excel_worksheet,
          lv_rangesheetname_old  type string,
          lv_rangesheetname_new  type string,
          lo_ranges_iterator     type ref to zcl_excel_collection_iterator,
          lo_range               type ref to zcl_excel_range,
          lv_range_value         type zif_excel_data_decl=>zexcel_range_value,
          lv_errormessage        type string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Check whether title consists only of allowed characters
* Illegal characters are: / \ [ ] * ? : --> http://msdn.microsoft.com/en-us/library/ff837411.aspx
* Illegal characters not in documentation:   ' as first character
*--------------------------------------------------------------------*
    if ip_title ca '/\[]*?:'.
      lv_errormessage = 'Found illegal character in sheetname. List of forbidden characters: /\[]*?:'(402).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if ip_title is not initial and ip_title(1) = `'`.
      lv_errormessage = 'Sheetname may not start with &'(403).   " & used instead of ' to allow fallbacklanguage
      replace '&' in lv_errormessage with `'`.
      zcx_excel=>raise_text( lv_errormessage ).
    endif.


*--------------------------------------------------------------------*
* Check whether title is unique in workbook
*--------------------------------------------------------------------*
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    while lo_worksheets_iterator->has_next( ) = abap_true.

      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      check me->guid <> lo_worksheet->get_guid( ).  " Don't check against itself
      if ip_title = lo_worksheet->get_title( ).  " Not unique --> raise exception
        lv_errormessage = 'Duplicate sheetname &'.
        replace '&' in lv_errormessage with ip_title.
        zcx_excel=>raise_text( lv_errormessage ).
      endif.

    endwhile.

*--------------------------------------------------------------------*
* Remember old sheetname and rename sheet to desired name
*--------------------------------------------------------------------*
    lv_rangesheetname_old = zcl_excel_common=>escape_string( me->title ) && '!'.
    me->title = ip_title.

*--------------------------------------------------------------------*
* After changing this worksheet's title we have to adjust
* all ranges that are referring to this worksheet.
*--------------------------------------------------------------------*
    lv_rangesheetname_new = zcl_excel_common=>escape_string( me->title ) && '!'.

    lo_ranges_iterator = me->excel->get_ranges_iterator( ).  "workbookglobal ranges
    while lo_ranges_iterator->has_next( ) = abap_true.

      lo_range ?= lo_ranges_iterator->get_next( ).
      lv_range_value = lo_range->get_value( ).
      replace all occurrences of lv_rangesheetname_old in lv_range_value with lv_rangesheetname_new.
      if sy-subrc = 0.
        lo_range->set_range_value( lv_range_value ).
      endif.

    endwhile.

    if me->ranges is bound.  "not yet bound if called from worksheet's constructor
      lo_ranges_iterator = me->get_ranges_iterator( ).  "sheetlocal ranges, repeat rows and columns
      while lo_ranges_iterator->has_next( ) = abap_true.

        lo_range ?= lo_ranges_iterator->get_next( ).
        lv_range_value = lo_range->get_value( ).
        replace all occurrences of lv_rangesheetname_old in lv_range_value with lv_rangesheetname_new.
        if sy-subrc = 0.
          lo_range->set_range_value( lv_range_value ).
        endif.

      endwhile.
    endif.


  endmethod.                    "SET_TITLE


  method update_dimension_range.

    data: ls_sheet_content type zif_excel_data_decl=>zexcel_s_cell_data,
          lv_row_alpha     type string,
          lv_column_alpha  type zif_excel_data_decl=>zexcel_cell_column_alpha.

    check sheet_content is not initial.

    upper_cell-cell_row = rows->get_min_index( ).
    if upper_cell-cell_row = 0.
      upper_cell-cell_row = zcl_excel_common=>c_excel_sheet_max_row.
    endif.
    upper_cell-cell_column = zcl_excel_common=>c_excel_sheet_max_col.

    lower_cell-cell_row = rows->get_max_index( ).
    if lower_cell-cell_row = 0.
      lower_cell-cell_row = zcl_excel_common=>c_excel_sheet_min_row.
    endif.
    lower_cell-cell_column = zcl_excel_common=>c_excel_sheet_min_col.

    loop at sheet_content into ls_sheet_content.
      if upper_cell-cell_row > ls_sheet_content-cell_row.
        upper_cell-cell_row = ls_sheet_content-cell_row.
      endif.
      if upper_cell-cell_column > ls_sheet_content-cell_column.
        upper_cell-cell_column = ls_sheet_content-cell_column.
      endif.
      if lower_cell-cell_row < ls_sheet_content-cell_row.
        lower_cell-cell_row = ls_sheet_content-cell_row.
      endif.
      if lower_cell-cell_column < ls_sheet_content-cell_column.
        lower_cell-cell_column = ls_sheet_content-cell_column.
      endif.
    endloop.

    upper_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = upper_cell-cell_column i_row = upper_cell-cell_row ).

    lower_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lower_cell-cell_column i_row = lower_cell-cell_row ).

  endmethod.                    "UPDATE_DIMENSION_RANGE


  method zif_excel_sheet_printsettings~clear_print_repeat_columns.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    clear:  me->print_title_col_from,
            me->print_title_col_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_COLUMNS


  method zif_excel_sheet_printsettings~clear_print_repeat_rows.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    clear:  me->print_title_row_from,
            me->print_title_row_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_ROWS


  method zif_excel_sheet_printsettings~get_print_repeat_columns.
    ev_columns_from = me->print_title_col_from.
    ev_columns_to   = me->print_title_col_to.
  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_COLUMNS


  method zif_excel_sheet_printsettings~get_print_repeat_rows.
    ev_rows_from = me->print_title_row_from.
    ev_rows_to   = me->print_title_row_to.
  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_ROWS


  method zif_excel_sheet_printsettings~set_print_repeat_columns.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    data: lv_col_from_int type i,
          lv_col_to_int   type i,
          lv_errormessage type string.


    lv_col_from_int = zcl_excel_common=>convert_column2int( iv_columns_from ).
    lv_col_to_int   = zcl_excel_common=>convert_column2int( iv_columns_to ).

*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    if lv_col_from_int < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if  lv_col_from_int > lv_col_to_int.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_col_from = iv_columns_from.
    me->print_title_col_to   = iv_columns_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).

  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_COLUMNS


  method zif_excel_sheet_printsettings~set_print_repeat_rows.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    data:     lv_errormessage                 type string.


*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    if iv_rows_from < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

    if  iv_rows_from > iv_rows_to.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    endif.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_row_from = iv_rows_from.
    me->print_title_row_to   = iv_rows_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  endmethod.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_ROWS


  method zif_excel_sheet_properties~get_right_to_left.
    result = right_to_left.
  endmethod.


  method zif_excel_sheet_properties~get_style.
    if zif_excel_sheet_properties~style is not initial.
      ep_style = zif_excel_sheet_properties~style.
    else.
      ep_style = me->excel->get_default_style( ).
    endif.
  endmethod.                    "ZIF_EXCEL_SHEET_PROPERTIES~GET_STYLE


  method zif_excel_sheet_properties~initialize.

    zif_excel_sheet_properties~show_zeros   = zif_excel_sheet_properties=>c_showzero.
    zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on.
    zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_on.

* inizialize zoomscale values
    zif_excel_sheet_properties~zoomscale = 100.
    zif_excel_sheet_properties~zoomscale_normal = 100.
    zif_excel_sheet_properties~zoomscale_pagelayoutview = 100 .
    zif_excel_sheet_properties~zoomscale_sheetlayoutview = 100 .
  endmethod.                    "ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE


  method zif_excel_sheet_properties~set_right_to_left.
    me->right_to_left = right_to_left.
  endmethod.


  method zif_excel_sheet_properties~set_style.
    zif_excel_sheet_properties~style = ip_style.
  endmethod.                    "ZIF_EXCEL_SHEET_PROPERTIES~SET_STYLE


  method zif_excel_sheet_protection~initialize.

    me->zif_excel_sheet_protection~protected = zif_excel_sheet_protection=>c_unprotected.
    clear me->zif_excel_sheet_protection~password.
    me->zif_excel_sheet_protection~auto_filter            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_cells           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_hyperlinks      = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~objects                = zif_excel_sheet_protection=>c_noactive.
*  me->zif_excel_sheet_protection~password               = zif_excel_sheet_protection=>c_noactive. "issue #68
    me->zif_excel_sheet_protection~pivot_tables           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~protected              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~scenarios              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_locked_cells    = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_unlocked_cells  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sheet                  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sort                   = zif_excel_sheet_protection=>c_noactive.

  endmethod.                    "ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE


  method zif_excel_sheet_vba_project~set_codename.
    me->zif_excel_sheet_vba_project~codename = ip_codename.
  endmethod.                    "ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME


  method zif_excel_sheet_vba_project~set_codename_pr.
    me->zif_excel_sheet_vba_project~codename_pr = ip_codename_pr.
  endmethod.                    "ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME_PR
endclass.
