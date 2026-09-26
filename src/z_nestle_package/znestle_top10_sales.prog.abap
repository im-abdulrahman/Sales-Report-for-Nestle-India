*&---------------------------------------------------------------------*
*& Report ZNESTLE_TOP10_SALES
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT znestle_top10_sales.

TABLES: vbak, vbap.

" Selection screen for date range filtering
SELECT-OPTIONS: s_erdat FOR vbak-erdat.

TYPES: BEGIN OF ty_sales,
         matnr     TYPE vbap-matnr,
         arktx     TYPE vbap-arktx,
         total_qty TYPE vbap-kwmeng,
         total_rev TYPE vbap-netwr,
         waers     TYPE vbak-waerk,
       END OF ty_sales,
       tt_sales TYPE STANDARD TABLE OF ty_sales.

DATA: gt_sales TYPE tt_sales,
      go_alv   TYPE REF TO cl_salv_table,
      gx_root  TYPE REF TO cx_root,
      gv_text  TYPE string.

START-OF-SELECTION.

  " Aggregate top 10 selling products joining VBAK and VBAP
  SELECT b~matnr,
         b~arktx,
         SUM( b~kwmeng ) AS total_qty,
         SUM( b~netwr ) AS total_rev,
         MAX( a~waerk ) AS waers
    FROM vbak AS a
    JOIN vbap AS b ON a~vbeln = b~vbeln
    WHERE a~erdat IN @s_erdat
    GROUP BY b~matnr, b~arktx
    ORDER BY total_rev DESCENDING
    INTO TABLE @gt_sales
    UP TO 10 ROWS.

  IF sy-subrc NE 0.
    MESSAGE 'No sales records found for the selected date range.' TYPE 'I' DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " Render Report using Object-Oriented ALV (CL_SALV_TABLE)
  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = go_alv
        CHANGING
          t_table      = gt_sales ).

      " Enable standard toolbar functions (Includes Excel download / Spreadsheet export)
      DATA(lo_functions) = go_alv->get_functions( ).
      lo_functions->set_all( abap_true ).

      " Optimize column display widths
      DATA(lo_columns) = go_alv->get_columns( ).
      lo_columns->set_optimize( abap_true ).

      " Display the Grid
      go_alv->display( ).

    CATCH cx_root INTO gx_root.
      gv_text = gx_root->get_text( ).
      MESSAGE gv_text TYPE 'E'.
  ENDTRY.
