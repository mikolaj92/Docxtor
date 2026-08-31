from __future__ import annotations

from docx.oxml.ns import qn

W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")
W_DEL_TEXT = qn("w:delText")
W_SDT = qn("w:sdt")
W_SDT_CONTENT = qn("w:sdtContent")
W_TBL = qn("w:tbl")
W_TXBX_CONTENT = qn("w:txbxContent")
V_TEXTBOX = "{urn:schemas-microsoft-com:vml}textbox"
WPS_TXBX = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}txbx"
R_ID = qn("r:id")
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"
_COMMENTS_IDS_PART = "word/commentsIds.xml"
_COMMENTS_EXTENSIBLE_PART = "word/commentsExtensible.xml"
_PEOPLE_PART = "word/people.xml"
_DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"
_CONTENT_TYPES_PART = "[Content_Types].xml"
_THREAD_PARTS = {
    _COMMENTS_EXTENDED_PART,
    _COMMENTS_IDS_PART,
    _COMMENTS_EXTENSIBLE_PART,
    _PEOPLE_PART,
}
_THREAD_REL_BY_TARGET = {
    "commentsExtended.xml": (
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "application/vnd.ms-word.commentsExtended+xml",
    ),
    "commentsIds.xml": (
        "http://schemas.microsoft.com/office/2016/relationships/commentsIds",
        "application/vnd.ms-word.commentsIds+xml",
    ),
    "commentsExtensible.xml": (
        "http://schemas.microsoft.com/office/2018/relationships/commentsExtensible",
        "application/vnd.ms-word.commentsExtensible+xml",
    ),
    "people.xml": (
        "http://schemas.microsoft.com/office/2011/relationships/people",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
    ),
}

_TEXT_NODE_TAGS = {W_T, W_DEL_TEXT}
_UNSUPPORTED_MOVE_TAGS = {
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
}
