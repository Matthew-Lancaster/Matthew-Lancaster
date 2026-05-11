Attribute VB_Name = "modkWabGetAll"
Option Explicit
Option Base 1

Public Sub GetAllTags(arry() As Long)
   Dim k As Long
   ReDim arry(1 To 1)
   k = 0
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ACKNOWLEDGEMENT_MODE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ALTERNATE_RECIPIENT_ALLOWED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AUTHORIZING_USERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AUTO_FORWARD_COMMENT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AUTO_FORWARD_COMMENT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AUTO_FORWARDED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_CORRELATOR
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_IDENTIFIER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_IDENTIFIER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_LENGTH
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_RETURN_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSATION_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSION_EITS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSION_WITH_LOSS_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERTED_EITS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEFERRED_DELIVERY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELIVER_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISCARD_REASON
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISCLOSURE_OF_RECIPIENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DL_EXPANSION_HISTORY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DL_EXPANSION_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_EXPIRY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IMPLICIT_CONVERSION_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IMPORTANCE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LATEST_DELIVERY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_CLASS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_CLASS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_DELIVERY_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_SECURITY_LABEL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OBSOLETED_IPMS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIPIENT_NAME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_EITS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_CERTIFICATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_RETURN_ADDRESS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PARENT_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIORITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGIN_CHECK
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROOF_OF_SUBMISSION_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_READ_RECEIPT_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIPT_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_REASSIGNMENT_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REDIRECTION_HISTORY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RELATED_IPMS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENSITIVITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LANGUAGES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LANGUAGES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPLY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_TAG
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RETURNED_IPM
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SECURITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_INCOMPLETE_COPY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENSITIVITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBJECT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBJECT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBJECT_IPM
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CLIENT_SUBMIT_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_X400_CONTENT_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBJECT_PREFIX_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBJECT_PREFIX_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NON_RECEIPT_REASON
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_READ_RECEIPT_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_SUBMISSION_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_SUBMIT_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SUBJECT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SUBJECT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISC_VAL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIG_MESSAGE_CLASS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIG_MESSAGE_CLASS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SUBMIT_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPLY_RECIPIENT_ENTRIES
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPLY_RECIPIENT_NAMES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPLY_RECIPIENT_NAMES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_READ_RECEIPT_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DELIVERY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_TO_ME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_CC_ME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_RECIP_ME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_START_DATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_END_DATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OWNER_APPT_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESPONSE_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENT_REPRESENTING_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENDER_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSATION_TOPIC_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSATION_TOPIC_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSATION_INDEX
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_BCC_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_BCC_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_CC_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_CC_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_TO_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_TO_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVED_BY_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RCVD_REPRESENTING_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSPORT_MESSAGE_HEADERS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSPORT_MESSAGE_HEADERS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELEGATION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TNEF_CORRELATION_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BODY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BODY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_TEXT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORT_TEXT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORTING_DL_NAME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPORTING_MTA_CERTIFICATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_BODY_CRC
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_BODY_COUNT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_BODY_TAG_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_BODY_TAG_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_COMPRESSED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_PREFIX_COUNT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_SYNC_TRAILING_COUNT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINALLY_INTENDED_RECIP_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_INTEGRITY_CHECK
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_EXPLICIT_CONVERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_RETURN_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_TOKEN
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NDR_REASON_CODE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NDR_DIAG_CODE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NON_RECEIPT_NOTIFICATION_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELIVERY_POINT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_DELIVERY_MODE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_DELIVERY_REPORT_REQUEST
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_FORWARDING_ADDRESS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_FORWARDING_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PHYSICAL_RENDITION_ATTRIBUTES
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROOF_OF_DELIVERY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROOF_OF_DELIVERY_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_CERTIFICATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_NUMBER_FOR_ADVICE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_NUMBER_FOR_ADVICE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REGISTERED_MAIL_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REPLY_REQUESTED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REQUESTED_DELIVERY_METHOD
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUPPLEMENTARY_INFO_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUPPLEMENTARY_INFO_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TYPE_OF_MTS_USER
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENDER_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CURRENT_VERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELETE_AFTER_SUBMIT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_BCC_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_BCC_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_CC_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_CC_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_TO_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_TO_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PARENT_DISPLAY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PARENT_DISPLAY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_DELIVERY_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_FLAGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_SIZE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PARENT_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SENTMAIL_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CORRELATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CORRELATE_MTSID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISCRETE_VALUES
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESPONSIBILITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SPOOLER_STATUS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSPORT_STATUS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_RECIPIENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_ATTACHMENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBMIT_FLAGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECIPIENT_STATUS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSPORT_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MSG_STATUS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MESSAGE_DOWNLOAD_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CREATION_VERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MODIFY_VERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HASATTACH
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BODY_CRC
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NORMALIZED_SUBJECT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NORMALIZED_SUBJECT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RTF_IN_SYNC
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_SIZE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_NUM
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PREPROCESS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINATING_MTA_CERTIFICATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROOF_OF_SUBMISSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OBJECT_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ICON
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MINI_ICON
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STORE_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STORE_RECORD_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECORD_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MAPPING_SIGNATURE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ACCESS_LEVEL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_INSTANCE_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ROW_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ACCESS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ROWID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ADDRTYPE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ADDRTYPE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_EMAIL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_EMAIL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMMENT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMMENT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEPTH
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_DISPLAY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_DISPLAY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CREATION_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LAST_MODIFICATION_TIME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESOURCE_FLAGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_DLL_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_DLL_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_UID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROVIDER_ORDINAL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_VERSION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_VERSION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CLSID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CONTACT_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CONTACT_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CATEGORY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CATEGORY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CATEGORY_SUB_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_CATEGORY_SUB_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_HOST_MAP
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_HIDDEN
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_DESIGNER_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_DESIGNER_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_DESIGNER_GUID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FORM_MESSAGE_BEHAVIOR
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEFAULT_STORE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STORE_SUPPORT_MASK
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STORE_STATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_SUBTREE_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_OUTBOX_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_WASTEBASKET_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_SENTMAIL_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MDB_PROVIDER
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RECEIVE_FOLDER_SETTINGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_VALID_FOLDER_MASK
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_SUBTREE_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_OUTBOX_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_WASTEBASKET_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IPM_SENTMAIL_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_VIEWS_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMMON_VIEWS_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FINDER_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_FLAGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FOLDER_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_COUNT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENT_UNREAD
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CREATE_TEMPLATES
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DETAILS_TABLE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SEARCH
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SELECTABLE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SUBFOLDERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATUS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ANR_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ANR_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTENTS_SORT_ORDER
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_HIERARCHY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_CONTENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FOLDER_ASSOCIATED_CONTENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEF_CREATE_DL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEF_CREATE_MAILUSER
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_CLASS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_CLASS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTAINER_MODIFY_VERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_PROVIDER_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEFAULT_VIEW_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ASSOC_CONTENT_COUNT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACHMENT_X400_PARAMETERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_DATA_OBJ
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_DATA_BIN
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_ENCODING
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_EXTENSION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_EXTENSION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_FILENAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_FILENAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_METHOD
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_LONG_FILENAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_LONG_FILENAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_PATHNAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_PATHNAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_RENDERING
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_TAG
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RENDERING_POSITION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_TRANSPORT_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_TRANSPORT_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_LONG_PATHNAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_LONG_PATHNAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_MIME_TAG_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_MIME_TAG_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ATTACH_ADDITIONAL_INFO
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TEMPLATEID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIMARY_CAPABILITY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_7BIT_DISPLAY_NAME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ACCOUNT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ACCOUNT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ALTERNATE_RECIPIENT
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CALLBACK_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CALLBACK_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONVERSION_PROHIBITED
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISCLOSE_RECIPIENTS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GENERATION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GENERATION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GIVEN_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GIVEN_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GOVERNMENT_ID_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GOVERNMENT_ID_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_INITIALS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_INITIALS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_KEYWORD_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_KEYWORD_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LANGUAGE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LANGUAGE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LOCATION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LOCATION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MAIL_PERMISSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MHS_COMMON_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MHS_COMMON_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORGANIZATIONAL_ID_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORGANIZATIONAL_ID_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SURNAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SURNAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_DISPLAY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ORIGINAL_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POSTAL_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POSTAL_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPANY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPANY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TITLE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TITLE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEPARTMENT_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEPARTMENT_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OFFICE_LOCATION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OFFICE_LOCATION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIMARY_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIMARY_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS2_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS2_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MOBILE_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MOBILE_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RADIO_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RADIO_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CAR_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CAR_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSMITABLE_DISPLAY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSMITABLE_DISPLAY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PAGER_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PAGER_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_USER_CERTIFICATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIMARY_FAX_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PRIMARY_FAX_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_FAX_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_FAX_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_FAX_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_FAX_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COUNTRY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COUNTRY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LOCALITY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_LOCALITY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATE_OR_PROVINCE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATE_OR_PROVINCE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STREET_ADDRESS_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STREET_ADDRESS_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POSTAL_CODE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POSTAL_CODE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POST_OFFICE_BOX_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_POST_OFFICE_BOX_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TELEX_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TELEX_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ISDN_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ISDN_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ASSISTANT_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ASSISTANT_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME2_TELEPHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME2_TELEPHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ASSISTANT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_ASSISTANT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SEND_RICH_INFO
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_WEDDING_ANNIVERSARY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BIRTHDAY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOBBIES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOBBIES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MIDDLE_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MIDDLE_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_NAME_PREFIX_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DISPLAY_NAME_PREFIX_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROFESSION_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROFESSION_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PREFERRED_BY_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PREFERRED_BY_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SPOUSE_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SPOUSE_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPUTER_NETWORK_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPUTER_NETWORK_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CUSTOMER_ID_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CUSTOMER_ID_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TTYTDD_PHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TTYTDD_PHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FTP_SITE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FTP_SITE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_GENDER
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MANAGER_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_MANAGER_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NICKNAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_NICKNAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PERSONAL_HOME_PAGE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PERSONAL_HOME_PAGE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_HOME_PAGE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_BUSINESS_HOME_PAGE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_VERSION
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_ENTRYIDS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_ADDRTYPES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_ADDRTYPES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_DEFAULT_ADDRESS_INDEX
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_EMAIL_ADDRESSES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTACT_EMAIL_ADDRESSES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPANY_MAIN_PHONE_NUMBER_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_COMPANY_MAIN_PHONE_NUMBER_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CHILDRENS_NAMES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CHILDRENS_NAMES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_CITY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_CITY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_COUNTRY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_COUNTRY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_POSTAL_CODE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_POSTAL_CODE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_STATE_OR_PROVINCE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_STATE_OR_PROVINCE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_STREET_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_STREET_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_POST_OFFICE_BOX_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HOME_ADDRESS_POST_OFFICE_BOX_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_CITY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_CITY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_COUNTRY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_COUNTRY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_POSTAL_CODE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_POSTAL_CODE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_STATE_OR_PROVINCE_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_STATE_OR_PROVINCE_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_STREET_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_STREET_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_POST_OFFICE_BOX_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OTHER_ADDRESS_POST_OFFICE_BOX_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STORE_PROVIDERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_PROVIDERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_TRANSPORT_PROVIDERS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DEFAULT_PROFILE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_SEARCH_PATH
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_DEFAULT_DIR
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_DEFAULT_PAB
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_FILTERING_HOOKS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_DLL_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_DLL_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_ENTRY_NAME
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_UID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_EXTRA_UIDS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICES
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_SUPPORT_FILES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_SUPPORT_FILES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_DELETE_FILES_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SERVICE_DELETE_FILES_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_AB_SEARCH_PATH_UPDATE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROFILE_NAME_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_PROFILE_NAME_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IDENTITY_DISPLAY_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IDENTITY_DISPLAY_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IDENTITY_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESOURCE_METHODS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESOURCE_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATUS_CODE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_IDENTITY_SEARCH_KEY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_OWN_STORE_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESOURCE_PATH_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_RESOURCE_PATH_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATUS_STRING_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_STATUS_STRING_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_X400_DEFERRED_DELIVERY_CANCEL
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_HEADER_FOLDER_ENTRYID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REMOTE_PROGRESS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REMOTE_PROGRESS_TEXT_W
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REMOTE_PROGRESS_TEXT_A
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_REMOTE_VALIDATE_OK
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTROL_FLAGS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTROL_STRUCTURE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTROL_TYPE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELTAX
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_DELTAY
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_XPOS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_YPOS
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_CONTROL_ID
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_INITIAL_DETAILS_PANE
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PR_SEND_INTERNET_ENCODING
End Sub
Public Sub GetAllTypes(arry() As Long)
   Dim k As Long
   ReDim arry(1 To 1)
   k = 0
   
   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_UNSPECIFIED

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_NULL

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_SHORT

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_LONG

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_FLOAT

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_DOUBLE

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_BOOLEAN

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_CURRENCY

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_APPTIME

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_SYSTIME

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_STRING8

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_BINARY

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_UNICODE

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_CLSID

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_LONGLONG

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_SHORT

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_LONG

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_FLOAT

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_DOUBLE

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_CURRENCY

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_APPTIME

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_SYSTIME

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_BINARY

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_STRING8

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_UNICODE

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_CLSID

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_MV_LONGLONG

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_ERROR

   k = k + 1
   ReDim Preserve arry(1 To k)
   arry(k) = PT_OBJECT
End Sub

