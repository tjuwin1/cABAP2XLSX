interface zif_excel_style_changer
  public .

  methods apply
    importing
      ip_worksheet   type ref to zcl_excel_worksheet
      ip_column      type simple
      ip_row         type zif_excel_data_decl=>zexcel_cell_row
    returning
      value(ep_guid) type zif_excel_data_decl=>zexcel_cell_style
    raising
      zcx_excel.
  methods get_guid
    returning
      value(result) type zif_excel_data_decl=>zexcel_cell_style.
  methods set_complete
    importing
      ip_complete   type zif_excel_data_decl=>zexcel_s_cstyle_complete
      ip_xcomplete  type zif_excel_data_decl=>zexcel_s_cstylex_complete
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_complete_font
    importing
      ip_font       type zif_excel_data_decl=>zexcel_s_cstyle_font
      ip_xfont      type zif_excel_data_decl=>zexcel_s_cstylex_font optional
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_complete_fill
    importing
      ip_fill       type zif_excel_data_decl=>zexcel_s_cstyle_fill
      ip_xfill      type zif_excel_data_decl=>zexcel_s_cstylex_fill optional
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_complete_borders
    importing
      ip_borders    type zif_excel_data_decl=>zexcel_s_cstyle_borders
      ip_xborders   type zif_excel_data_decl=>zexcel_s_cstylex_borders optional
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_complete_alignment
    importing
      ip_alignment  type zif_excel_data_decl=>zexcel_s_cstyle_alignment
      ip_xalignment type zif_excel_data_decl=>zexcel_s_cstylex_alignment optional
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_complete_protection
    importing
      ip_protection  type zif_excel_data_decl=>zexcel_s_cstyle_protection
      ip_xprotection type zif_excel_data_decl=>zexcel_s_cstylex_protection optional
    returning
      value(result)  type ref to zif_excel_style_changer.
  methods set_complete_borders_all
    importing
      ip_borders_allborders  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_allborders type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)          type ref to zif_excel_style_changer.
  methods set_complete_borders_diagonal
    importing
      ip_borders_diagonal  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_diagonal type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)        type ref to zif_excel_style_changer.
  methods set_complete_borders_down
    importing
      ip_borders_down  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_down type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)    type ref to zif_excel_style_changer.
  methods set_complete_borders_left
    importing
      ip_borders_left  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_left type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)    type ref to zif_excel_style_changer.
  methods set_complete_borders_right
    importing
      ip_borders_right  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_right type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)     type ref to zif_excel_style_changer.
  methods set_complete_borders_top
    importing
      ip_borders_top  type zif_excel_data_decl=>zexcel_s_cstyle_border
      ip_xborders_top type zif_excel_data_decl=>zexcel_s_cstylex_border optional
    returning
      value(result)   type ref to zif_excel_style_changer.
  methods set_number_format
    importing
      value         type zif_excel_data_decl=>zexcel_number_format
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_bold
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_color_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_family
    importing
      value         type zif_excel_data_decl=>zexcel_style_font_family
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_italic
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_name
    importing
      value         type zif_excel_data_decl=>zexcel_style_font_name
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_scheme
    importing
      value         type zif_excel_data_decl=>zexcel_style_font_scheme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_size
    importing
      value         type numeric
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_strikethrough
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_underline
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_font_underline_mode
    importing
      value         type zif_excel_data_decl=>zexcel_style_font_underline
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_filltype
    importing
      value         type zif_excel_data_decl=>zexcel_fill_type
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_rotation
    importing
      value         type zif_excel_data_decl=>zexcel_rotation
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_fgcolor
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_fgcolor_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_fgcolor_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_fgcolor_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_fgcolor_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_bgcolor
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_bgcolor_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_bgcolor_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_bgcolor_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_bgcolor_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_type
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-type
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_degree
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-degree
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_bottom
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-bottom
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_left
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-left
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_top
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-top
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_right
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-right
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_position1
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-position1
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_position2
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-position2
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_fill_gradtype_position3
    importing
      value         type zif_excel_data_decl=>zexcel_s_gradient_type-position3
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_mode
    importing
      value         type zif_excel_data_decl=>zexcel_diagonal
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_horizontal
    importing
      value         type zif_excel_data_decl=>zexcel_alignment
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_vertical
    importing
      value         type zif_excel_data_decl=>zexcel_alignment
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_textrotation
    importing
      value         type zif_excel_data_decl=>zexcel_text_rotation
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_wraptext
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_shrinktofit
    importing
      value         type abap_boolean
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_alignment_indent
    importing
      value         type zif_excel_data_decl=>zexcel_indent
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_protection_hidden
    importing
      value         type zif_excel_data_decl=>zexcel_cell_protection
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_protection_locked
    importing
      value         type zif_excel_data_decl=>zexcel_cell_protection
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allborders_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allborders_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allbo_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allbo_color_indexe
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allbo_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_allbo_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_color_ind
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_color_the
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_diagonal_color_tin
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_color_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_down_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_color_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_left_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_color_indexe
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_right_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_style
    importing
      value         type zif_excel_data_decl=>zexcel_border
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_color
    importing
      value         type zif_excel_data_decl=>zexcel_s_style_color
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_color_rgb
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_argb
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_color_indexed
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_indexed
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_color_theme
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_theme
    returning
      value(result) type ref to zif_excel_style_changer.
  methods set_borders_top_color_tint
    importing
      value         type zif_excel_data_decl=>zexcel_style_color_tint
    returning
      value(result) type ref to zif_excel_style_changer.
  data: complete_style  type zif_excel_data_decl=>zexcel_s_cstyle_complete read-only,
        complete_stylex type zif_excel_data_decl=>zexcel_s_cstylex_complete read-only.
endinterface.
