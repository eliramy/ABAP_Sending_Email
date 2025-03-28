CLASS y_send_mail DEFINITION
  PUBLIC
  CREATE PUBLIC.

  PUBLIC SECTION.

*"* public components of class Y_SEND_MAIL
*"* do not include other source files here!!!
    CLASS-METHODS send_mail
      IMPORTING
        !iv_sender_user      TYPE sy-uname DEFAULT sy-uname
        !iv_sender_email     TYPE ad_smtpadr DEFAULT ''
        !iv_doc_type         TYPE so_obj_tp DEFAULT 'RAW'
        !iv_message_subject  TYPE so_obj_des
        !it_message_body     TYPE soli_tab
        !it_recipient        TYPE rmps_recipient_bcs
        !iv_send_at          TYPE bcs_sndat OPTIONAL
        !it_itcoo            TYPE yyt_itcoo OPTIONAL
        !it_tline            TYPE yyt_tline OPTIONAL
        !it_file             TYPE yyt_file OPTIONAL
        !iv_commit           TYPE xfeld DEFAULT 'X'
        !iv_send_immediately TYPE os_boolean DEFAULT 'X'.
    CLASS-METHODS tab_to_excel_xml
      IMPORTING
        !ir_data TYPE REF TO data
      CHANGING
        !cv_size TYPE i
        !ct_xml  TYPE solix_tab.
    CLASS-METHODS convert_spool_otf
      IMPORTING
        !iv_rqident    TYPE tsp01-rqident
      RETURNING
        VALUE(et_file) TYPE solix_tab.
    CLASS-METHODS convert_str_to_file
      IMPORTING
        !iv_str              TYPE string
      RETURNING
        VALUE(et_file_solix) TYPE solix_tab.
*        !is_stamp         TYPE yys_app_stamp
    CLASS-METHODS send_appointment
      IMPORTING
        !iv_organizer     TYPE syuname
        !iv_location      TYPE string
        !iv_title         TYPE sc_txtshor
        !it_body_txt      TYPE so_txttab
        !it_email         TYPE bcsy_smtpa
        !iv_all_day_event TYPE xfeld
        !iv_app_type      TYPE sc_appttyp DEFAULT 'MEETING'
      EXPORTING
        !ev_app_class     TYPE seoclsname
        !ev_guid          TYPE sc_aptguid.
    CLASS-METHODS cancel_appointment
      IMPORTING
        !lv_app_class TYPE seoclsname
        !iv_guid      TYPE sc_aptguid
      EXCEPTIONS
        no_appointment_found.
    CLASS-METHODS convert_spool_adsp
      IMPORTING
        !iv_rqident    TYPE tsp01-rqident
      RETURNING
        VALUE(et_file) TYPE solix_tab.
  PROTECTED SECTION.
*"* protected components of class Y_SEND_MAIL
*"* do not include other source files here!!!
  PRIVATE SECTION.

*"* private components of class Y_SEND_MAIL
*"* do not include other source files here!!!
    CLASS-DATA lo_send_request TYPE REF TO cl_bcs.
    CLASS-DATA lo_document TYPE REF TO cl_document_bcs.

    CLASS-METHODS send_request.
    CLASS-METHODS body_subject
      IMPORTING
        !iv_doc_type        TYPE so_obj_tp
        !iv_message_subject TYPE so_obj_des
        !it_message_body    TYPE soli_tab.
    CLASS-METHODS attach_itcoo
      IMPORTING
        !it_otf TYPE yyt_itcoo.
    CLASS-METHODS attach_text_cont
      IMPORTING
        !iv_attach_subject TYPE so_obj_des
        !iv_type           TYPE so_obj_tp
        !it_att_text       TYPE soli_tab.
    CLASS-METHODS attach_tline
      IMPORTING
        !it_otf TYPE yyt_tline.
    CLASS-METHODS attach_file
      IMPORTING
        !it_bin_file TYPE yyt_file.
    CLASS-METHODS attach_bin_cont
      IMPORTING
        !iv_attach_subject TYPE so_obj_des
        !iv_type           TYPE so_obj_tp
        !it_att_bin        TYPE solix_tab
        !iv_size           TYPE so_obj_len.
    CLASS-METHODS create_sender
      IMPORTING
        !iv_sender       TYPE sy-uname
        !iv_sender_email TYPE ad_smtpadr.
    CLASS-METHODS create_recipients
      IMPORTING
        !it_recipient TYPE rmps_recipient_bcs.
    CLASS-METHODS execute_send
      IMPORTING
        !iv_commit           TYPE xfeld OPTIONAL
        !iv_send_at          TYPE bcs_sndat
        !iv_send_immediately TYPE os_boolean.
    CLASS-METHODS cancel_attachment_ics
      IMPORTING
        !iv_organizer        TYPE adr6-smtp_addr
        !io_appointment      TYPE REF TO cl_appointment
      RETURNING
        VALUE(lt_attachment) TYPE soli_tab.
ENDCLASS.



CLASS y_send_mail IMPLEMENTATION.


  METHOD attach_bin_cont.

    DATA: lx_document_bcs TYPE REF TO cx_document_bcs VALUE IS INITIAL.

    TRY.
        lo_document->add_attachment(
        EXPORTING
        i_attachment_type    = iv_type
        i_attachment_subject = iv_attach_subject
        i_attachment_size    = iv_size
        i_att_content_hex    = it_att_bin ) .

      CATCH cx_document_bcs INTO lx_document_bcs.

    ENDTRY.

  ENDMETHOD.


  METHOD attach_file.

    DATA:
    lt_objfile TYPE solix_tab.

    FIELD-SYMBOLS:
    <fs_bin_file> TYPE yys_file.

    LOOP AT it_bin_file ASSIGNING <fs_bin_file>.
      CLEAR lt_objfile.

      APPEND LINES OF <fs_bin_file>-t_file TO lt_objfile.

      CALL METHOD attach_bin_cont
        EXPORTING
          iv_attach_subject = <fs_bin_file>-attachment_subject
          iv_type           = <fs_bin_file>-attachment_type
          it_att_bin        = lt_objfile
          iv_size           = <fs_bin_file>-attachment_size.

    ENDLOOP.

  ENDMETHOD.


  METHOD attach_itcoo.

    DATA:
      lt_objbin         TYPE soli_tab,
      lt_tline          TYPE TABLE OF tline,
      lv_len_in         TYPE sood-objlen,
      lt_doctab_archive TYPE TABLE OF docs,
      lt_record         TYPE srm_t_solisti1.

    FIELD-SYMBOLS:
    <fs_otf> TYPE yys_itcoo.

    LOOP AT it_otf ASSIGNING <fs_otf>.

      CLEAR: lv_len_in,
             lt_doctab_archive,
             lt_tline,
             lt_record,
             lt_objbin.

      CALL FUNCTION 'CONVERT_OTF_2_PDF'
        EXPORTING
          use_otf_mc_cmd         = 'X'
        IMPORTING
          bin_filesize           = lv_len_in
        TABLES
          otf                    = <fs_otf>-t_itcoo
          doctab_archive         = lt_doctab_archive
          lines                  = lt_tline
        EXCEPTIONS
          err_conv_not_possible  = 1
          err_otf_mc_noendmarker = 2
          OTHERS                 = 3.

      IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.

      CALL FUNCTION 'QCE1_CONVERT'
        TABLES
          t_source_tab         = lt_tline
          t_target_tab         = lt_record
        EXCEPTIONS
          convert_not_possible = 1
          OTHERS               = 2.

      IF sy-subrc <> 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      lt_objbin[] = lt_record[].

      CALL METHOD attach_text_cont
        EXPORTING
          iv_attach_subject = <fs_otf>-attachment_subject
          iv_type           = <fs_otf>-attachment_type
          it_att_text       = lt_objbin.

    ENDLOOP.

  ENDMETHOD.


  METHOD attach_text_cont.

    DATA: lx_document_bcs TYPE REF TO cx_document_bcs VALUE IS INITIAL.

    TRY.
        lo_document->add_attachment(
        EXPORTING
        i_attachment_type    = iv_type
        i_attachment_subject = iv_attach_subject
        i_att_content_text   = it_att_text ) .

      CATCH cx_document_bcs INTO lx_document_bcs.

    ENDTRY.

  ENDMETHOD.


  METHOD attach_tline.

    DATA:
    lt_objbin TYPE soli_tab.

    FIELD-SYMBOLS:
    <fs_otf> TYPE yys_tline.

    LOOP AT it_otf ASSIGNING <fs_otf>.

      CLEAR: lt_objbin.

      CALL FUNCTION 'QCE1_CONVERT'
        TABLES
          t_source_tab         = <fs_otf>-t_tline
          t_target_tab         = lt_objbin
        EXCEPTIONS
          convert_not_possible = 1
          OTHERS               = 2.

      IF sy-subrc <> 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      CALL METHOD attach_text_cont
        EXPORTING
          iv_attach_subject = <fs_otf>-attachment_subject
          iv_type           = <fs_otf>-attachment_type
          it_att_text       = lt_objbin.

    ENDLOOP.

  ENDMETHOD.


  METHOD body_subject.

    lo_document =
     cl_document_bcs=>create_document( i_type    = iv_doc_type   "'RAW'
                                       i_text    = it_message_body
                                       i_subject = iv_message_subject ).
  ENDMETHOD.


  METHOD cancel_appointment.

*  FIELD-SYMBOLS:
*  <fs_o_appointment> TYPE REF TO cl_appointment,
*  <fs_participants>  TYPE scspart,
*  <fs_container>     TYPE swcont.
*
*  DATA:
*  lo_document        TYPE REF TO cl_document_bcs,
*  lo_receiver        TYPE REF TO if_recipient_bcs,
*  lo_email           TYPE REF TO cl_bcs,
*  lt_attachment      TYPE soli_tab,
*  lt_appointments    TYPE screftab,
*  lv_guid             TYPE sc_aptguid,
*  lv_found           TYPE xfeld,
*  lt_app_guids       TYPE scappidtab,
*  lt_participants    TYPE scparttab,
*  lv_object          TYPE swotrtime-object,
*  lt_container       TYPE swconttab,
*  lv_address         TYPE adr6-smtp_addr,
*  lv_send_result(1)  TYPE c VALUE IS INITIAL,
*  lv_persnumber      TYPE usr21-persnumber,
*  lv_addrnumber      TYPE usr21-addrnumber,
*  lv_organizer       TYPE syuname.
*
*  lo_email = cl_bcs=>create_persistent( ).
*  lt_appointments = cl_appointment=>select_by_application_guids( application_guids = lt_app_guids
*                                                                 appointment_class = lv_app_class ).
*
*  LOOP AT lt_appointments ASSIGNING <fs_o_appointment>.
*    CHECK iv_guid = <fs_o_appointment>->get_guid( ).
*    lv_found = 'X'.
*    EXIT.
*  ENDLOOP.
*
*  IF lv_found IS INITIAL.
*    RAISE no_appointment_found.
*  ENDIF.
*
** Find participants
*  CALL METHOD <fs_o_appointment>->get_participants
*    IMPORTING
*      participants          = lt_participants
*    EXCEPTIONS
*      no_participants       = 1
*      appointment_not_exist = 2
*      OTHERS                = 3.
*
*  IF sy-subrc <> 0.
**   MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
**              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
*  ENDIF.
*
*  LOOP AT lt_participants ASSIGNING <fs_participants>.
*
*    CALL FUNCTION 'SWO_CREATE'
*      EXPORTING
*        objtype           = <fs_participants>-objtype
*        objkey            = <fs_participants>-objkey
*      IMPORTING
*        object            = lv_object
*      EXCEPTIONS
*        no_remote_objects = 1
*        OTHERS            = 2.
*
*    IF sy-subrc <> 0.
** Implement suitable error handling here
*    ENDIF.
*
*    CALL FUNCTION 'SWO_INVOKE'
*      EXPORTING
*        access    = 'G'
*        object    = lv_object
*        verb      = 'AddressString'
*      TABLES
*        container = lt_container.
*
*    READ TABLE lt_container ASSIGNING <fs_container> INDEX 1.
*
*    lv_address = <fs_container>-value.
*    lo_receiver = cl_cam_address_bcs=>create_internet_address( lv_address ).
*    lo_email->add_recipient( i_recipient = lo_receiver ).
*
*  ENDLOOP.
*
** Send Notifcation also to organizer
*  lv_organizer = <fs_o_appointment>->get_organizer( ).
*
*  SELECT SINGLE b~smtp_addr INTO lv_address
*  FROM usr21 AS a INNER JOIN adr6 AS b
*  ON ( a~persnumber = b~persnumber
*  AND  a~addrnumber = b~addrnumber )
*  WHERE a~bname = lv_organizer.
*
*  IF sy-subrc = 0.
*    lo_receiver = cl_cam_address_bcs=>create_internet_address( lv_address ).
*    lo_email->add_recipient( i_recipient = lo_receiver ).
*  ENDIF.
*
*  lt_attachment = cancel_attachment_ics(  iv_organizer   = lv_address
*                                          io_appointment = <fs_o_appointment> ).
*
*  lo_document = cl_document_bcs=>create_from_text( i_text         = lt_attachment
*                                                   i_documenttype = 'ICS'
*                                                   i_subject      = 'Cancellation' ).
*  lo_email->set_document( lo_document ).
*
*  lo_email->set_send_immediately( 'X' ).
*
*  lo_email->send( EXPORTING i_with_error_screen = ''
*                  RECEIVING result = lv_send_result ).
*
*  CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
*    EXPORTING
*      wait = 'X'.


  ENDMETHOD.


  METHOD cancel_attachment_ics.

*  DATA:
*  lv_row       TYPE so_text255,
*  lv_cnt_nn    TYPE n LENGTH 2,
*  lv_blank_tt  TYPE c LENGTH 2 VALUE '',
*  lv_date_from TYPE sc_datefro,
*  lv_time_from TYPE sc_timefro,
*  lv_date_to   TYPE sc_dateto,
*  lv_time_to   TYPE sc_timefro.
*
*  CALL METHOD io_appointment->get_date
*    IMPORTING
*      date_from = lv_date_from
*      time_from = lv_time_from
*      date_to   = lv_date_to
*      time_to   = lv_time_to.
*
*  lv_row = 'BEGIN:VCALENDAR'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'VERSION:2.0'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'PRODID:-//SAP AG//R/3-702//E'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'METHOD:CANCEL'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'BEGIN:VEVENT'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  "lv_row = 'UID:E009E0E7344BA2F1BEDD005056AC611E@saphwdf.sap.corp'.
*  DATA:
*  lv_appt_guid     TYPE sc_aptguid,
*  lv_appt_guid_str TYPE string.
*
*  CALL METHOD io_appointment->get_guid
*    RECEIVING
*      guid = lv_appt_guid.
*
*  lv_appt_guid_str = lv_appt_guid.
*
*  "append the e-mail address domain to the GUID!
*  CONCATENATE lv_appt_guid_str '@menora.co.il' INTO lv_appt_guid_str.
*
*  lv_row = 'UID:'.
*  CONCATENATE lv_row lv_appt_guid_str INTO lv_row.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'SEQUENCE:1'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'DTSTAMP:'.
*  CONCATENATE lv_row sy-datum 'T' sy-uzeit 'Z' INTO lv_row.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  CONCATENATE 'ORGANIZER:' iv_organizer INTO lv_row.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  " lv_row = 'DTSTART:20101231T110000Z'.
*  lv_row = 'DTSTART:'.
*  CONCATENATE lv_row lv_date_from 'T' lv_time_from 'Z' INTO lv_row.
*  lv_cnt_nn = strlen( lv_row ) .
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'DTEND:'.
*  CONCATENATE lv_row lv_date_to 'T' lv_time_to 'Z' INTO lv_row.
*  lv_cnt_nn = strlen( lv_row ) .
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'SUMMARY: Cancellation'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'STATUS:CANCELLED'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*  lv_row = 'END:VEVENT'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  lv_row = 'END:VCALENDAR'.
*  lv_cnt_nn = strlen( lv_row ).
*  CONCATENATE '0' lv_cnt_nn lv_blank_tt lv_row INTO lv_row RESPECTING BLANKS.
*  APPEND lv_row TO lt_attachment.
*
*  CALL FUNCTION 'SO_RAW_INT_TO_RTF'
*    TABLES
*      objcont_old = lt_attachment
*      objcont_new = lt_attachment.

  ENDMETHOD.


  METHOD convert_spool_adsp.

    FIELD-SYMBOLS:
    <fs_partlist> TYPE adspartdesc.

    DATA:
      ls_tsp01    TYPE tsp01sys,
      lv_data     TYPE fpcontent,
      lt_partlist TYPE TABLE OF adspartdesc.

    CALL FUNCTION 'RSPO_ISELECT_TSP01'
      EXPORTING
        rfcsystem  = sy-sysid
        rqident    = iv_rqident
      IMPORTING
        tsp01_elem = ls_tsp01
      EXCEPTIONS
        error      = 1
        OTHERS     = 2.

    IF sy-subrc <> 0.
* Implement suitable error handling here
    ENDIF.

    CALL FUNCTION 'RSPO_ADSP_FILL_PARTLIST'
      EXPORTING
        rq       = ls_tsp01
      TABLES
        partlist = lt_partlist.

    READ TABLE lt_partlist ASSIGNING <fs_partlist> INDEX 1.

    CHECK sy-subrc = 0.

    CALL FUNCTION 'FPCOMP_CREATE_PDF_FROM_SPOOL'
      EXPORTING
        i_spoolid      = iv_rqident
        i_partnum      = <fs_partlist>-adsnum
      IMPORTING
        e_pdf          = lv_data
      EXCEPTIONS
        ads_error      = 1
        usage_error    = 2
        system_error   = 3
        internal_error = 4
        OTHERS         = 5.

    IF sy-subrc <> 0.
* Implement suitable error handling here
    ENDIF.

    CHECK lv_data IS NOT INITIAL.

    FIELD-SYMBOLS:
    <fs_file_x> TYPE cps_x255.

    DATA:
      lt_file_x TYPE cpt_x255,
      lv_size   TYPE i,
      ls_file   TYPE solix.

    CALL METHOD cl_scp_change_db=>xstr_to_xtab
      EXPORTING
        im_xstring = lv_data
      IMPORTING
        ex_xtab    = lt_file_x
        ex_size    = lv_size.

    LOOP AT lt_file_x ASSIGNING <fs_file_x>.
      ls_file-line = <fs_file_x>.
      APPEND ls_file TO et_file.
    ENDLOOP.

  ENDMETHOD.


  METHOD convert_spool_otf.

    DATA:
    lt_soli TYPE TABLE OF soli.

    CALL FUNCTION 'RSPO_RETURN_SPOOLJOB'
      EXPORTING
        rqident              = iv_rqident
        desired_type         = 'OTF'
      TABLES
        buffer               = lt_soli
      EXCEPTIONS
        no_such_job          = 1
        job_contains_no_data = 2
        selection_empty      = 3
        no_permission        = 4
        can_not_access       = 5
        read_error           = 6
        type_no_match        = 7
        OTHERS               = 8.

    IF sy-subrc <> 0.
    ENDIF.

    CHECK lt_soli[] IS NOT INITIAL.

    DATA:
      lv_size  TYPE i,
      lv_data  TYPE xstring,
      lv_dummy TYPE soli_tab.

* CONVERT OTF TO PDF
    CALL FUNCTION 'CONVERT_OTF'
      EXPORTING
        format                = 'PDF'
      IMPORTING
        bin_filesize          = lv_size
        bin_file              = lv_data
      TABLES
        otf                   = lt_soli[]
        lines                 = lv_dummy
      EXCEPTIONS
        err_max_linewidth     = 1
        err_format            = 2
        err_conv_not_possible = 3
        OTHERS                = 4.

    IF sy-subrc <> 0.
*    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*    WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    CHECK lv_data IS NOT INITIAL.

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer     = lv_data
      TABLES
        binary_tab = et_file.

  ENDMETHOD.


  METHOD convert_str_to_file.

    FIELD-SYMBOLS:
    <fs_tab>     TYPE lxe_xtab.

    DATA:
      lv_xstr       TYPE xstring,
      ls_file_solix TYPE solix,
* lt_file_soli  TYPE soli_tab,
      lt_tab        TYPE TABLE OF lxe_xtab.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text     = iv_str
        encoding = '1824'     " UTF-8
      IMPORTING
        buffer   = lv_xstr
      EXCEPTIONS
        failed   = 1
        OTHERS   = 2.

    IF sy-subrc <> 0.
* Implement suitable error handling here
    ENDIF.

    CALL FUNCTION 'LXE_COMMON_XSTRING_TO_TABLE'
      EXPORTING
        in_xstring = lv_xstr
      TABLES
        ex_tab     = lt_tab.

    LOOP AT lt_tab ASSIGNING <fs_tab>.
      ls_file_solix-line = <fs_tab>-text.
      APPEND ls_file_solix TO et_file_solix.
    ENDLOOP.

  ENDMETHOD.


  METHOD create_recipients.

    FIELD-SYMBOLS: <fs_recipient> TYPE REF TO if_recipient_bcs.

    LOOP AT it_recipient ASSIGNING <fs_recipient>.
* Set recipient
      lo_send_request->add_recipient( EXPORTING
                                         i_recipient = <fs_recipient>
                                         i_express   = '' ).
    ENDLOOP.

  ENDMETHOD.


  METHOD create_sender.

    DATA:
    lo_sender	TYPE REF TO	if_sender_bcs.

    IF iv_sender_email IS INITIAL.
      lo_sender = cl_sapuser_bcs=>create( iv_sender ).
    ELSE.
      lo_sender =
      cl_cam_address_bcs=>create_internet_address( iv_sender_email ).
    ENDIF.

* Set sender
    lo_send_request->set_sender( EXPORTING
                                 i_sender = lo_sender ).

  ENDMETHOD.


  METHOD execute_send.

    DATA: lv_sent_to_all(1) TYPE c VALUE IS INITIAL.

    IF iv_send_at IS INITIAL.
      lo_send_request->set_send_immediately( iv_send_immediately ).
    ELSE.
      lo_send_request->send_request->set_send_at( iv_send_at ).
      lo_send_request->set_send_immediately( space ).
    ENDIF.

    lo_send_request->set_status_attributes(
        EXPORTING
          i_requested_status = 'N'
          i_status_mail      = 'N' ).

    lo_send_request->send(
      EXPORTING
        i_with_error_screen = ''
      RECEIVING
        result              = lv_sent_to_all ).

    CHECK iv_commit = 'X'.
    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        wait = 'X'.

  ENDMETHOD.


  METHOD send_appointment.

*  TYPE-POOLS: sccon.
*
*  DATA: swo_objid  TYPE swotobjid,
*        swo_return TYPE swotreturn.
*
*  FIELD-SYMBOLS: <fs_email> TYPE ad_smtpadr.
*
*  DATA lo_appointment       TYPE REF TO cl_appointment.
*  DATA ls_participant       TYPE scspart.
*  DATA lv_address           TYPE obj_record.
*  DATA ls_address_container TYPE TABLE OF swcont.
*
*  DATA lo_send_request      TYPE REF TO cl_bcs.
*  DATA lv_sent_to_all       TYPE os_boolean.
*
*  CREATE OBJECT lo_appointment.
*
** Add multiple attendees
*  LOOP AT it_email ASSIGNING <fs_email>.
*    CLEAR ls_participant.
*
** swc_create_object lv_address 'ADDRESS' space.
*
*    swo_objid-objtype = 'ADDRESS'.
*    swo_objid-objkey  = space.
*
*    lv_address-header = 'OBJH'.
*    lv_address-type   = 'SWO '.
*
*    CALL FUNCTION 'SWO_CREATE'
*      EXPORTING
*        objtype = swo_objid-objtype
*        objkey  = swo_objid-objkey
*      IMPORTING
*        object  = lv_address-handle
*        return  = swo_return.
*
*    IF swo_return-code NE 0.
*      lv_address-handle = 0.
*    ENDIF.
*
** swc_set_element ls_address_container 'AddressString' email-low.
*    CALL FUNCTION 'SWC_ELEMENT_SET'
*      EXPORTING
*        element       = 'AddressString'
*        field         = <fs_email>
*      TABLES
*        container     = ls_address_container
*      EXCEPTIONS
*        type_conflict = 1
*        OTHERS        = 2.
*
*    IF sy-subrc <> 0.
** Implement suitable error handling here
*    ENDIF.
*
**   swc_set_element ls_address_container 'TypeId' 'U'.
*    CALL FUNCTION 'SWC_ELEMENT_SET'
*      EXPORTING
*        element       = 'TypeId'
*        field         = 'U'
*      TABLES
*        container     = ls_address_container
*      EXCEPTIONS
*        type_conflict = 1
*        OTHERS        = 2.
*
*    IF sy-subrc <> 0.
** Implement suitable error handling here
*    ENDIF.
*
*
**   swc_call_method lv_address 'Create' ls_address_container.
*    CALL FUNCTION 'SWO_INVOKE'
*      EXPORTING
*        access     = 'C'
*        object     = lv_address-handle
*        verb       = 'Create'
*        persistent = ' '
*      IMPORTING
*        return     = swo_return
*      TABLES
*        container  = ls_address_container.
*
*    IF swo_return-code NE 0.
*      sy-msgid = swo_return-workarea.
*      sy-msgno = swo_return-message.
*      sy-msgty = 'E'.
*      sy-msgv1 = swo_return-variable1.
*      sy-msgv2 = swo_return-variable2.
*      sy-msgv3 = swo_return-variable3.
*      sy-msgv4 = swo_return-variable4.
*    ENDIF.
*    sy-subrc = swo_return-code.
*
*    CHECK sy-subrc = 0.
** * get key and type of object
**    swc_get_object_key lv_address ls_participant-objkey.
**    swc_get_object_type lv_address ls_participant-objtype.
*
*    CALL FUNCTION 'SWO_OBJECT_ID_GET'
*      EXPORTING
*        object = lv_address-handle
*      IMPORTING
*        return = swo_return
*        objid  = swo_objid.
*
*    IF swo_return-code NE 0.
*      sy-msgid = swo_return-workarea.
*      sy-msgno = swo_return-message.
*      sy-msgty = 'E'.
*      sy-msgv1 = swo_return-variable1.
*      sy-msgv2 = swo_return-variable2.
*      sy-msgv3 = swo_return-variable3.
*      sy-msgv4 = swo_return-variable4.
*    ENDIF.
*    sy-subrc = swo_return-code.
*    ls_participant-objkey  = swo_objid-objkey.
*    ls_participant-objtype = swo_objid-objtype.
*
*    CHECK sy-subrc = 0.
*
*    MOVE sccon_part_sndmail_with_ans TO ls_participant-send_mail.
*    ls_participant-comm_mode = 'INT'.
*    lo_appointment->add_participant( participant = ls_participant ).
*
*  ENDLOOP.
*
*
** Apppointment for specific date/time
*  lo_appointment->set_date( date_from = is_stamp-date_from
*                            time_from = is_stamp-time_from
*                            date_to   = is_stamp-date_to
*                            time_to   = is_stamp-time_to
*                            timezone  = is_stamp-timezone ).
*
** Make appointment appear "busy"
*  lo_appointment->set_busy_value( sccon_busy_busy ).
*
*
** Set Location
*  lo_appointment->set_location_string( iv_location ).
*
*
** Set Organizer
*  lo_appointment->set_organizer( iv_organizer ).
*
** "Type of meeting" (value picked from table SCAPPTTYPE)
*  lo_appointment->set_type( iv_app_type ).
*
** Make this an all day event
*  IF iv_all_day_event = 'X'.
*    lo_appointment->set_view_attributes( show_on_top = iv_all_day_event ).
*  ENDIF.
*
** Set Meeting body text
*  lo_appointment->set_text( it_body_txt ).
*
** Set Meeting Subject
*  lo_appointment->set_title( iv_title ).
*
** Important to set this one to space. Otherwise SAP will send a not user-friendly e-mail
*  lo_appointment->save( send_invitation = space ).
*
** Return App. Class
*  ev_app_class = lo_appointment->get_appointment_class( ).
*
** Return GUID
*  ev_guid = lo_appointment->get_guid( ).
*
** Now that we have the appointment, we can send a good one for outlook by switching to BCS
*  lo_send_request = lo_appointment->create_send_request( ).
*
*  lo_send_request->set_send_immediately( 'X' ).
*
** don't request read/delivery receipts
*  lo_send_request->set_status_attributes(
*  i_requested_status = 'N' i_status_mail = 'N' ).
*
** Send it to the world
*  lv_sent_to_all = lo_send_request->send( i_with_error_screen = 'X' ).
*
*  CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
*    EXPORTING
*      wait = 'X'.

  ENDMETHOD.


  METHOD send_mail.

    CLEAR: lo_send_request, lo_document.

* Create send request
    CALL METHOD send_request.

* Message body and subject
    CALL METHOD body_subject
      EXPORTING
        iv_doc_type        = iv_doc_type
        iv_message_subject = iv_message_subject
        it_message_body    = it_message_body.

* Add attachment
    IF it_itcoo[] IS NOT INITIAL.
      CALL METHOD attach_itcoo
        EXPORTING
          it_otf = it_itcoo.
    ENDIF.

    IF it_tline[] IS NOT INITIAL.
      CALL METHOD attach_tline
        EXPORTING
          it_otf = it_tline.
    ENDIF.

    IF it_file[] IS NOT INITIAL.
      CALL METHOD attach_file
        EXPORTING
          it_bin_file = it_file.
    ENDIF.

* Pass the document to send request
    lo_send_request->set_document( lo_document ).

* Create & set sender
    CALL METHOD create_sender
      EXPORTING
        iv_sender       = iv_sender_user
        iv_sender_email = iv_sender_email.

* Create recipients
    CALL METHOD create_recipients
      EXPORTING
        it_recipient = it_recipient.

* Execute Send email
    CALL METHOD execute_send
      EXPORTING
        iv_commit           = iv_commit
        iv_send_at          = iv_send_at
        iv_send_immediately = iv_send_immediately.

  ENDMETHOD.


  METHOD send_request.

    lo_send_request = cl_bcs=>create_persistent( ).

  ENDMETHOD.


  METHOD tab_to_excel_xml.

    FIELD-SYMBOLS:
    <fs_tab> TYPE ANY TABLE.

    ASSIGN ir_data->* TO <fs_tab>.

    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table   = DATA(lo_salv_table) " Basis Class Simple ALV Tables
                                CHANGING  t_table        = <fs_tab> ).         " Internal table to display
      CATCH cx_salv_msg.

    ENDTRY.

* Convert internal table to xml (format xlsx)
    DATA(lv_xml) = lo_salv_table->to_xml( xml_type = if_salv_bs_xml=>c_type_xlsx ).

* Size xml
    cv_size = xstrlen( lv_xml ).

* Convert xml xstring to internal table solix_tab
    ct_xml = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xml ).

  ENDMETHOD.
ENDCLASS.
