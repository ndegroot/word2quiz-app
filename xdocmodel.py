""" not used for now: easy recognizing the type of the lists  not solved"""
# import docx
#
#
# # noinspection PyProtectedMember
# def iter_paragraphs(parent, recursive=True):
#     """
#     Yield each paragraph and table child within *parent*, in document order.
#     Each returned value is an instance of Paragraph. *parent*
#     would most commonly be a reference to a main Document object, but
#     also works for a _Cell object, which itself can contain paragraphs and tables.
#     """
#     if isinstance(parent, docx.document.Document):
#         parent_elm = parent.element.body
#     elif isinstance(parent, docx.table._Cell):
#         parent_elm = parent._tc
#     else:
#         raise TypeError(repr(type(parent)))
#
#     for child in parent_elm.iterchildren():
#         if isinstance(child, docx.oxml.text.paragraph.CT_P):
#             yield docx.text.paragraph.Paragraph(child, parent)
#         elif isinstance(child, docx.oxml.table.CT_Tbl):
#             if recursive:
#                 table = docx.table.Table(child, parent)
#                 for row in table.rows:
#                     for cell in row.cells:
#                         for child_paragraph in iter_paragraphs(cell):
#                             yield child_paragraph
