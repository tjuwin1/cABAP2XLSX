interface zif_excel_data_decl
  public.
  types zexcel_active_worksheet        type i.
  types zexcel_aes_password            type c length 50.
  types zexcel_alignment               type c length 20.
  types zexcel_application             type c length 80.
  types zexcel_border                  type c length 20.
  types zexcel_break                   type i.
  types zexcel_category                type c length 80.
  types zexcel_cell_column             type i.
  types zexcel_cell_column_alpha       type c length 3.
  types zexcel_cell_coords             type string.
  types zexcel_cell_data_type          type string.
  types zexcel_cell_formula            type string.
  types zexcel_cell_protection         type c length 1. "( 1 or 0 )
  types zexcel_cell_row                type i.
  types zexcel_cell_style              type xstring.
  types zexcel_cell_value              type string.
  types zexcel_color                   type c length 8.
  types zexcel_column_name             type c length 255.
  types zexcel_company                 type c length 80.
  types zexcel_conditional_type        type c length 10.
  types zexcel_conditional_value       type string.
  types zexcel_condition_operator      type c length 20.
  types zexcel_condition_rule          type c length 20.
  types zexcel_condition_rule_iconset  type c length 20.
  types zexcel_converter_option_filter type c length 1. "( ` `, X, -)
  types zexcel_converter_option_hidehd type c length 1.
  types zexcel_converter_option_hidenc type c length 1.
  types zexcel_converter_option_subtot type c length 1.
  types zexcel_creator                 type c length 80.
  types zexcel_data_val_error_style    type c length 20.
  types zexcel_data_val_operator       type c length 20.
  types zexcel_data_val_type           type c length 20.
  types zexcel_dec_8_2                 type p length 8 decimals 2.
  types zexcel_description             type c length 80.
  types zexcel_diagonal                type i.
  types zexcel_docsecurity             type n length 1.
  types zexcel_drawing_anchor          type c length 3. " ONE, ABS, TWO
  types zexcel_drawing_type            type c length 5. " image, chart
  types zexcel_fieldname               type c length 30.
  types zexcel_fill_type               type c length 20.
  types zexcel_graph_type              type i.
  types zexcel_guid                    type xstring.
  types zexcel_hidden                  type c length 1.
  types zexcel_indent                  type i.
  types zexcel_keywords                type c length 80.
  types zexcel_locked                  type c length 1.
  types zexcel_number_format           type string.
  types zexcel_pane_state              type string.
  types zexcel_pane_type               type string.
  types zexcel_print_gridlines         type string.
  types zexcel_pwd_hash                type xstring.
  types zexcel_range_guid              type xstring.
  types zexcel_range_name              type string.
  types zexcel_range_value             type string.
  types zexcel_revisionspassword       type c length 80.
  types zexcel_rotation                type i.
  types zexcel_scalecrop               type abap_boolean.
  types zexcel_sheet_hidden            type abap_boolean.
  types zexcel_sheet_orienatation      type c length 20.
  types zexcel_sheet_paper_size        type i.
  types zexcel_sheet_selected          type abap_boolean.
  types zexcel_sheet_showzeros         type c length 1.
  types zexcel_sheet_summary           type i.
  types zexcel_sheet_title             type string.
  types zexcel_sheet_zoomscale         type i.
  types zexcel_show_gridlines          type abap_boolean.
  types zexcel_show_rowcolheader       type abap_boolean.
  types zexcel_style_color_argb        type c length 8.
  types zexcel_style_color_component   type c length 2.
  types zexcel_style_color_indexed     type i.
  types zexcel_style_color_theme       type i.
  types zexcel_style_color_tint        type f.
  types zexcel_style_font_family       type i.
  types zexcel_style_font_name         type string.
  types zexcel_style_font_size         type i.
  types zexcel_style_font_scheme       type c length 20.
  types zexcel_style_font_underline    type c length 20.
  types zexcel_style_formula           type string.
  types zexcel_style_priority          type i.
  types zexcel_subject                 type c length 80.
  types zexcel_table_style             type string.
  types zexcel_table_totals_function   type string.
  types zexcel_text_rotation           type i.
  types zexcel_title                   type string.
  types zexcel_validation_formula1     type string.
  types zexcel_workbookpassword        type c length 80.
  types zexcel_worksheets_name         type c length 80.
  types zexcel_column_id               type i.
  types zexcel_component_position      type n length 4.
  types zexcel_convexit                type c length 5.
  types zexcel_disp_text_long          type c length 40.
  types zexcel_disp_text_medium        type c length 20.
  types zexcel_disp_text_short         type c length 10.
  types zexcel_key_color_override      type c length 1.
  types zexcel_screen_display          type abap_boolean.

  types: begin of zexcel_conditional_above_avg,
           above_average      type xsdboolean,
           equal_average      type xsdboolean,
           standard_deviation type n length 1,
           cell_style         type zexcel_cell_style,
         end of  zexcel_conditional_above_avg.

  types: begin of zexcel_conditional_cellis,
           formula    type zexcel_style_formula,
           formula2   type zexcel_style_formula,
           operator   type zexcel_condition_operator,
           cell_style type zexcel_cell_style,
         end of zexcel_conditional_cellis.

  types: begin of zexcel_conditional_colorscale,
           cfvo1_type  type zexcel_conditional_type,
           cfvo1_value type zexcel_conditional_value,
           cfvo2_type  type zexcel_conditional_type,
           cfvo2_value type zexcel_conditional_value,
           cfvo3_type  type zexcel_conditional_type,
           cfvo3_value type zexcel_conditional_value,
           colorrgb1   type zexcel_color,
           colorrgb2   type zexcel_color,
           colorrgb3   type zexcel_color,
         end of zexcel_conditional_colorscale.

  types: begin of zexcel_conditional_databar,
           cfvo1_type  type zexcel_conditional_type,
           cfvo1_value type zexcel_conditional_value,
           cfvo2_type  type zexcel_conditional_type,
           cfvo2_value type zexcel_conditional_value,
           colorrgb    type zexcel_color,
         end of zexcel_conditional_databar.

  types: begin of zexcel_conditional_expression,
           formula    type zexcel_style_formula,
           cell_style type zexcel_cell_style,
         end of zexcel_conditional_expression.

  types: begin of zexcel_conditional_iconset,
           iconset     type zexcel_condition_rule_iconset,
           cfvo1_type  type zexcel_conditional_type,
           cfvo1_value type zexcel_conditional_value,
           cfvo2_type  type zexcel_conditional_type,
           cfvo2_value type zexcel_conditional_value,
           cfvo3_type  type zexcel_conditional_type,
           cfvo3_value type zexcel_conditional_value,
           cfvo4_type  type zexcel_conditional_type,
           cfvo4_value type zexcel_conditional_value,
           cfvo5_type  type zexcel_conditional_type,
           cfvo5_value type zexcel_conditional_value,
           showvalue   type c length 1, " Condition type
         end of zexcel_conditional_iconset.

  types: begin of zexcel_conditional_top10,
           topxx_count type i,
           percent     type abap_boolean,
           bottom      type abap_boolean,
           cell_style  type zexcel_cell_style,
         end of zexcel_conditional_top10.

  types: begin of zexcel_drawing_location,
           col        type int4,
           col_offset type int4,
           row        type int4,
           row_offset type int4,
         end of zexcel_drawing_location.

  types: begin of zexcel_drawing_size,
           width  type int4,
           height type int4,
         end of zexcel_drawing_size.

  types: begin of zexcel_drawing_position,
           anchor type zexcel_drawing_anchor,
           from   type zexcel_drawing_location,
           to     type zexcel_drawing_location,
           size   type zexcel_drawing_size,
         end of zexcel_drawing_position.

  types: begin of zexcel_pane,
           ysplit      type zexcel_cell_row,
           xsplit      type zexcel_cell_row,
           topleftcell type zexcel_cell_coords,
           activepane  type zexcel_pane_type,
           state       type zexcel_pane_state,
         end of zexcel_pane.

  types: begin of zexcel_s_autofilter_area,
           row_start type zexcel_cell_row,
           col_start type zexcel_cell_column,
           row_end   type zexcel_cell_row,
           col_end   type zexcel_cell_column,
         end of zexcel_s_autofilter_area.

  types: begin of zexcel_s_autofilter_values,
           column type zexcel_cell_column,
           value  type zexcel_cell_value,
         end of zexcel_s_autofilter_values.

  types: begin of zexcel_s_cellxfs,
           numfmtid          type int4,
           fontid            type int4,
           fillid            type int4,
           borderid          type int4,
           xfid              type int4,
           alignmentid       type int4,
           protectionid      type int4,
           applynumberformat type int4,
           applyfont         type int4,
           applyfill         type int4,
           applyborder       type int4,
           applyalignment    type int4,
           applyprotection   type int4,
         end of zexcel_s_cellxfs.

  types: begin of zexcel_s_style_color,
           rgb     type zexcel_style_color_argb,
           indexed type zexcel_style_color_indexed,
           theme   type zexcel_style_color_theme,
           tint    type zexcel_style_color_tint,
         end of zexcel_s_style_color.

  types: begin of zexcel_s_style_font,
           bold           type abap_boolean,
           italic         type abap_boolean,
           underline      type abap_boolean,
           underline_mode type c length 20,
           strikethrough  type abap_boolean,
           size           type i,
           color          type zexcel_s_style_color,
           name           type string,
           family         type i,
           scheme         type zexcel_style_font_scheme,
         end of zexcel_s_style_font.

  types: begin of zexcel_s_rtf,
           offset type i,
           length type i,
           font   type zexcel_s_style_font,
         end of zexcel_s_rtf,

         zexcel_t_rtf type sorted table of ZEXCEL_s_RTF with unique key offset.

  types: begin of zexcel_s_cell_data,
           cell_row          type zexcel_cell_row,
           cell_column       type zexcel_cell_column,
           cell_value        type zexcel_cell_value,
           cell_formula      type zexcel_cell_formula,
           cell_coords       type zexcel_cell_coords,
           cell_style        type zexcel_cell_style,
           data_type         type zexcel_cell_data_type,
           column_formula_id type i,
           rtf_tab           type zexcel_t_rtf,
           table             type ref to zcl_excel_table,
           table_fieldname   type zexcel_fieldname,
           table_header      type abap_boolean,
         end of zexcel_s_cell_data.

  types: begin of zexcel_s_converter_fil,
           rownumber  type zexcel_cell_row,
           columnname type zexcel_fieldname,
         end of zexcel_s_converter_fil.

  types: begin of zexcel_s_converter_layo,
           is_stripped        type abap_boolean,
           is_fixed           type abap_boolean,
           max_subtotal_level type i,
         end of zexcel_s_converter_layo.

  types: begin of zexcel_s_converter_option,
           filter           type zexcel_converter_option_filter,
           subtot           type zexcel_converter_option_subtot,
           hidenc           type zexcel_converter_option_hidenc,
           hidehd           type zexcel_converter_option_hidehd,
         end of zexcel_s_converter_option.

  types: begin of zexcel_s_cstylex_alignment,
           horizontal   type abap_boolean,
           vertical     type abap_boolean,
           textrotation type abap_boolean,
           wraptext     type abap_boolean,
           shrinktofit  type abap_boolean,
           indent       type abap_boolean,
         end of zexcel_s_cstylex_alignment.

  types: begin of zexcel_s_cstylex_color,
           rgb     type abap_boolean,
           indexed type abap_boolean,
           theme   type abap_boolean,
           tint    type abap_boolean,
         end of zexcel_s_cstylex_color.

  types: begin of zexcel_s_cstylex_border,
           border_style type abap_boolean,
           border_color type zexcel_s_cstylex_color,
         end of zexcel_s_cstylex_border.

  types: begin of zexcel_s_cstylex_borders,
           allborders    type zexcel_s_cstylex_border,
           diagonal      type zexcel_s_cstylex_border,
           diagonal_mode type abap_boolean,
           down          type zexcel_s_cstylex_border,
           left          type zexcel_s_cstylex_border,
           right         type zexcel_s_cstylex_border,
           top           type zexcel_s_cstylex_border,
         end of zexcel_s_cstylex_borders.

  types: begin of zexcel_s_cstylex_font,
           bold           type abap_boolean,
           color          type zexcel_s_cstylex_color,
           family         type abap_boolean,
           italic         type abap_boolean,
           name           type abap_boolean,
           scheme         type abap_boolean,
           size           type abap_boolean,
           strikethrough  type abap_boolean,
           underline      type abap_boolean,
           underline_mode type abap_boolean,
         end of zexcel_s_cstylex_font.

  types: begin of zexcel_s_cstylex_gradtype,
           type      type abap_boolean,
           degree    type abap_boolean,
           bottom    type abap_boolean,
           left      type abap_boolean,
           top       type abap_boolean,
           right     type abap_boolean,
           position1 type abap_boolean,
           position2 type abap_boolean,
           position3 type abap_boolean,
         end of zexcel_s_cstylex_gradtype.

  types: begin of zexcel_s_cstylex_fill,
           filltype type abap_boolean,
           rotation type abap_boolean,
           fgcolor  type zexcel_s_cstylex_color,
           bgcolor  type zexcel_s_cstylex_color,
           gradtype type zexcel_s_cstylex_gradtype,
         end of zexcel_s_cstylex_fill.

  types: begin of zexcel_s_cstylex_number_format,
           format_code type abap_boolean,
         end of zexcel_s_cstylex_number_format.

  types: begin of zexcel_s_cstylex_protection,
           hidden type abap_boolean,
           locked type abap_boolean,
         end of zexcel_s_cstylex_protection.

  types: begin of zexcel_s_cstylex_complete,
           font          type zexcel_s_cstylex_font,
           fill          type zexcel_s_cstylex_fill,
           borders       type zexcel_s_cstylex_borders,
           alignment     type zexcel_s_cstylex_alignment,
           number_format type zexcel_s_cstylex_number_format,
           protection    type zexcel_s_cstylex_protection,
         end of zexcel_s_cstylex_complete.

  types: begin of zexcel_s_cstyle_alignment,
           horizontal   type zexcel_alignment,
           vertical     type zexcel_alignment,
           textrotation type zexcel_text_rotation,
           wraptext     type abap_boolean,
           shrinktofit  type abap_boolean,
           indent       type zexcel_indent,
         end of zexcel_s_cstyle_alignment.

  types: begin of zexcel_s_cstyle_border,
           border_style type zexcel_border,
           border_color type zexcel_s_style_color,
         end of zexcel_s_cstyle_border.

  types: begin of zexcel_s_cstyle_font,
           bold           type abap_boolean,
           color          type zexcel_s_style_color,
           family         type zexcel_style_font_family,
           italic         type abap_boolean,
           name           type zexcel_style_font_name,
           scheme         type zexcel_style_font_scheme,
           size           type zexcel_style_font_size,
           strikethrough  type abap_boolean,
           underline      type abap_boolean,
           underline_mode type zexcel_style_font_underline,
         end of zexcel_s_cstyle_font.

  types: begin of zexcel_s_cstyle_borders,
           allborders    type zexcel_s_cstyle_border,
           diagonal      type zexcel_s_cstyle_border,
           diagonal_mode type zexcel_diagonal,
           down          type zexcel_s_cstyle_border,
           left          type zexcel_s_cstyle_border,
           right         type zexcel_s_cstyle_border,
           top           type zexcel_s_cstyle_border,
         end of zexcel_s_cstyle_borders.

  types: begin of zexcel_s_gradient_type,
           type      type zexcel_fill_type,
           degree    type c length 3,
           bottom    type c length 3,
           left      type c length 3,
           top       type c length 3,
           right     type c length 3,
           position1 type c length 3,
           position2 type c length 3,
           position3 type c length 3,
         end of zexcel_s_gradient_type.

  types: begin of zexcel_s_cstyle_fill,
           filltype type zexcel_fill_type,
           rotation type zexcel_rotation,
           fgcolor  type zexcel_s_style_color,
           bgcolor  type zexcel_s_style_color,
           gradtype type zexcel_s_gradient_type,
         end of zexcel_s_cstyle_fill.

  types: begin of zexcel_s_cstyle_number_format,
           format_code type zexcel_number_format,
         end of zexcel_s_cstyle_number_format.

  types: begin of zexcel_s_cstyle_protection,
           hidden type zexcel_cell_protection,
           locked type zexcel_cell_protection,
         end of zexcel_s_cstyle_protection.

  types: begin of zexcel_s_cstyle_complete,
           font          type zexcel_s_cstyle_font,
           fill          type zexcel_s_cstyle_fill,
           borders       type zexcel_s_cstyle_borders,
           alignment     type zexcel_s_cstyle_alignment,
           number_format type zexcel_s_cstyle_number_format,
           protection    type zexcel_s_cstyle_protection,
         end of zexcel_s_cstyle_complete.

  types: begin of zexcel_s_drawings,
           drawing type ref to zcl_excel_drawing,
         end of zexcel_s_drawings.

  types: begin of zexcel_s_fieldcatalog,
           tabname          type tabname,
           fieldname        type zexcel_fieldname,
           position         type zexcel_component_position,
           scrtext_l        type zexcel_disp_text_long,
           style            type zexcel_cell_style,
           style_header     type zexcel_cell_style,
           style_total      type zexcel_cell_style,
           style_cond       type zexcel_cell_style,
           totals_function  type zexcel_table_totals_function,
           formula          type abap_boolean,
           abap_type        type c length 1,
           column_formula   type string,
           column_name      type zexcel_column_name,
           currency_column  type zexcel_fieldname,
           unit_column      type zexcel_fieldname,
           text_column      type zexcel_fieldname,
           width            type int4,
         end of zexcel_s_fieldcatalog.

  types: begin of zexcel_s_shared_string,
           string_no    type int4,
           string_value type zexcel_cell_value,
           string_type  type zexcel_cell_data_type,
           rtf_tab      type zexcel_t_rtf,
         end of zexcel_s_shared_string.

  types: begin of zexcel_s_stylemapping,
           dynamic_style_guid type zexcel_cell_style,
           complete_style     type zexcel_s_cstyle_complete,
           complete_stylex    type zexcel_s_cstylex_complete,
           guid               type zexcel_cell_style,
           added_to_iterator  type abap_boolean,
         end of zexcel_s_stylemapping.

  types: begin of zexcel_s_styles_cond_mapping,
           guid  type xstring,
           style type int4,
           dxf   type int4,
         end of zexcel_s_styles_cond_mapping.

  types: begin of zexcel_s_styles_mapping,
           guid  type xstring,
           style type int4,
         end of zexcel_s_styles_mapping.

  types: begin of zexcel_s_style_alignment,
           horizontal   type zexcel_alignment,
           vertical     type zexcel_alignment,
           textrotation type zexcel_text_rotation,
           wraptext     type abap_boolean,
           shrinktofit  type abap_boolean,
           indent       type zexcel_indent,
         end of zexcel_s_style_alignment.

  types: begin of zexcel_s_style_border,
           left_color     type zexcel_s_style_color,
           left_style     type zexcel_border,
           right_color    type zexcel_s_style_color,
           right_style    type zexcel_border,
           top_color      type zexcel_s_style_color,
           top_style      type zexcel_border,
           bottom_color   type zexcel_s_style_color,
           bottom_style   type zexcel_border,
           diagonal_color type zexcel_s_style_color,
           diagonal_style type zexcel_border,
           diagonalup     type int1,
           diagonaldown   type int1,
         end of zexcel_s_style_border.

  types: begin of zexcel_s_style_fill,
           filltype type zexcel_fill_type,
           rotation type zexcel_rotation,
           fgcolor  type zexcel_s_style_color,
           bgcolor  type zexcel_s_style_color,
           gradtype type zexcel_s_gradient_type,
         end of zexcel_s_style_fill.


  types: begin of zexcel_s_style_numfmt,
           numfmt type zexcel_number_format,
         end of zexcel_s_style_numfmt.

  types: begin of zexcel_s_style_protection,
           locked type zexcel_locked,
           hidden type zexcel_hidden,
         end of zexcel_s_style_protection.

  types: begin of zexcel_s_tabcolor,
           rgb type zexcel_style_color_argb,
         end of zexcel_s_tabcolor.

  types: begin of zexcel_s_table_settings,
           table_style         type zexcel_table_style,
           table_name          type string,
           top_left_column     type zexcel_cell_column_alpha,
           top_left_row        type zexcel_cell_row,
           show_row_stripes    type abap_boolean,
           show_column_stripes type abap_boolean,
           bottom_right_column type zexcel_cell_column_alpha,
           bottom_right_row    type zexcel_cell_row,
           nofilters           type abap_boolean,
         end of zexcel_s_table_settings.

  types: begin of zexcel_s_worksheet_head_foot,
           left_value   type string,
           left_font    type zexcel_s_style_font,
           center_value type string,
           center_font  type zexcel_s_style_font,
           right_value  type string,
           right_font   type zexcel_s_style_font,
           left_image   type ref to zcl_excel_drawing,
           right_image  type ref to zcl_excel_drawing,
           center_image type ref to zcl_excel_drawing,
         end of zexcel_s_worksheet_head_foot.
  types: begin of zexcel_s_converter_col,
           rownumber  type zexcel_cell_row,
           columnname type zexcel_fieldname,
           fontcolor  type zexcel_style_color_argb,
           fillcolor  type zexcel_style_color_argb,
           nokeycol   type zexcel_key_color_override,
         end of zexcel_s_converter_col.

  types: begin of zexcel_s_converter_fcat,
           tabname         type tabname,
           fieldname       type zexcel_fieldname,
           columnname      type zexcel_fieldname,
           position        type zexcel_component_position,
           inttype         type c length 1,
           decimals        type int1,
           scrtext_s       type zexcel_disp_text_short,
           scrtext_m       type zexcel_disp_text_medium,
           scrtext_l       type zexcel_disp_text_long,
           totals_function type zexcel_table_totals_function,
           fix_column      type abap_boolean,
           alignment       type zexcel_alignment,
           is_optimized    type abap_boolean,
           is_hidden       type abap_boolean,
           is_collapsed    type abap_boolean,
           is_subtotalled  type abap_boolean,
           sort_level      type int4,
           style_hdr       type zexcel_cell_style,
           style_normal    type zexcel_cell_style,
           style_stripped  type zexcel_cell_style,
           style_total     type zexcel_cell_style,
           style_subtotal  type zexcel_cell_style,
           col_id          type zexcel_column_id,
           convexit        type zexcel_convexit,
         end of zexcel_s_converter_fcat.

  types zexcel_t_converter_col       type hashed table of ZEXCEL_s_CONVERTER_COL with unique key rownumber columnname.
  types zexcel_t_converter_fcat      type standard table of ZEXCEL_s_CONVERTER_FCAT with non-unique default key.

  types zexcel_t_autofilter_values   type standard table of ZEXCEL_s_AUTOFILTER_VALUES with non-unique default key.
  types zexcel_t_cellxfs             type standard table of ZEXCEL_s_CELLXFS with non-unique key
                numfmtid fontid fillid borderid xfid alignmentid protectionid applynumberformat applyfont applyfill applyborder applyalignment applyprotection.
  types zexcel_t_cell_data           type sorted table of ZEXCEL_s_CELL_DATA with unique key cell_row cell_column.
  types zexcel_t_cell_data_unsorted  type standard table of ZEXCEL_s_CELL_DATA with non-unique default key.
  types zexcel_t_converter_fil       type hashed table of ZEXCEL_s_CONVERTER_FIL with unique key rownumber columnname.
  types zexcel_t_drawings            type standard table of ZEXCEL_s_DRAWINGS with non-unique default key.
  types zexcel_t_fieldcatalog        type standard table of ZEXCEL_s_FIELDCATALOG with non-unique default key.
  types zexcel_t_shared_string       type sorted table of ZEXCEL_s_SHARED_STRING with non-unique key string_value.
  types zexcel_t_stylemapping1       type hashed table of ZEXCEL_s_STYLEMAPPING
      with unique key primary_key        components dynamic_style_guid complete_stylex complete_style
      with non-unique sorted key added_to_iterator components added_to_iterator guid.
  types zexcel_t_stylemapping2       type hashed table of ZEXCEL_s_STYLEMAPPING with unique key guid.
  types zexcel_t_styles_cond_mapping type standard table of ZEXCEL_s_STYLES_COND_MAPPING with non-unique default key.
  types zexcel_t_styles_mapping      type standard table of ZEXCEL_s_STYLES_MAPPING with non-unique default key.
  types zexcel_t_style_alignment     type standard table of ZEXCEL_s_STYLE_ALIGNMENT with non-unique key horizontal vertical textrotation wraptext shrinktofit indent.
  types zexcel_t_style_border        type standard table of ZEXCEL_s_STYLE_BORDER with non-unique key
                left_color left_style right_color right_style top_color top_style bottom_color bottom_style diagonal_color diagonal_style diagonalup diagonaldown.
  types zexcel_t_style_color_argb    type standard table of zexcel_style_color_argb with non-unique default key.
  types zexcel_t_style_fill          type standard table of ZEXCEL_s_STYLE_FILL with non-unique key filltype rotation fgcolor bgcolor gradtype.
  types zexcel_t_style_font          type standard table of ZEXCEL_s_STYLE_FONT with non-unique key bold italic underline UNDERLINE_mode strikethrough size color
        name family scheme.
  types zexcel_t_style_numfmt        type standard table of ZEXCEL_s_STYLE_NUMFMT with non-unique key numfmt.
  types zexcel_t_style_protection    type standard table of ZEXCEL_s_STYLE_PROTECTION with non-unique key locked hidden.

  types: begin of ty_syst,
           msgv1 type msgv1,
           msgv2 type msgv2,
           msgv3 type msgv3,
           msgv4 type msgv4,
           msgid type symsgid,
           msgno type symsgno,
           msgty type msgty,
         end of ty_syst.
endinterface.
