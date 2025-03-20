INTERFACE yif_prepare_mail
  PUBLIC .


  METHODS get_subject
    RETURNING
      VALUE(rv_subject) TYPE so_obj_des .
  METHODS get_body
    RETURNING
      VALUE(rt_body) TYPE soli_tab .
  METHODS get_distribution_list
    IMPORTING
      !iv_dliname         TYPE so_obj_nam
    RETURNING
      VALUE(rt_recipient) TYPE rmps_recipient_bcs .
  METHODS get_attach_file
    IMPORTING
      !ir_data       TYPE REF TO data
    RETURNING
      VALUE(rt_file) TYPE yyt_file .
ENDINTERFACE.
