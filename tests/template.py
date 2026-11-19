""" Synthetic template with the content control shapes of the real coupon (`Impressão POS 58mm v2.docm`) """

import docx
from docx.oxml import parse_xml

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
NS = f'xmlns:w="{W}" xmlns:w15="{W15}"'


def run(text: str) -> str:
    return f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def cell(text: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p>{run(text)}</w:p></w:tc>'


# `w:sdtContent` with `w:r`: run-level control inside a paragraph, placeholder split in two runs
CLIENT_NAME = f'''<w:p {NS}><w:r><w:t xml:space="preserve">Cliente: </w:t></w:r>
<w:sdt><w:sdtPr><w:id w:val="1"/><w:tag w:val="clientname"/><w:text/></w:sdtPr>
<w:sdtContent>{run("Nome")}{run(" do cliente")}</w:sdtContent></w:sdt></w:p>'''

# `w:sdtContent` with `w:p`: block-level control
TOTAL = f'''<w:sdt {NS}><w:sdtPr><w:id w:val="2"/><w:tag w:val="total"/></w:sdtPr>
<w:sdtContent><w:p>{run("0")}{run(",00")}</w:p><w:p>{run("sobra")}</w:p></w:sdtContent></w:sdt>'''

# `w:dropDownList`
PAYMENT_METHOD = f'''<w:p {NS}><w:sdt><w:sdtPr><w:id w:val="3"/><w:tag w:val="payment_method"/>
<w:dropDownList>
<w:listItem w:displayText="Escolha um item." w:value=""/>
<w:listItem w:displayText="Crédito via internet" w:value="credit_online"/>
<w:listItem w:displayText="Dinheiro" w:value="cash"/>
</w:dropDownList></w:sdtPr>
<w:sdtContent>{run("Escolha um item.")}</w:sdtContent></w:sdt></w:p>'''

# `w15:repeatingSection` around the rows of a table; one item with three cells
ORDER_ITEMS = f'''<w:tbl {NS}><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
<w:sdt><w:sdtPr><w:id w:val="4"/><w:tag w:val="order_items"/><w15:repeatingSection/></w:sdtPr>
<w:sdtContent><w:sdt><w:sdtPr><w:id w:val="5"/><w15:repeatingSectionItem/></w:sdtPr>
<w:sdtContent><w:tr>{cell("0x")}{cell("ITEM")}
<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:sdt><w:sdtPr><w:id w:val="6"/><w:tag w:val="item_price"/></w:sdtPr>
<w:sdtContent>{run("0,00")}</w:sdtContent></w:sdt></w:p></w:tc>
</w:tr></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:tbl>'''

# composite control: a table plus a paragraph without runs (the `return_change` of the real coupon)
RETURN_CHANGE = f'''<w:sdt {NS}><w:sdtPr><w:id w:val="7"/><w:tag w:val="return_change"/></w:sdtPr>
<w:sdtContent><w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>{cell("0,00")}</w:tr></w:tbl><w:p/></w:sdtContent></w:sdt>'''

# block-level control with an empty paragraph
NOTES = f'''<w:sdt {NS}><w:sdtPr><w:id w:val="8"/><w:tag w:val="notes"/></w:sdtPr>
<w:sdtContent><w:p/></w:sdtContent></w:sdt>'''

TAGS = ['clientname', 'total', 'payment_method', 'order_items', 'return_change', 'notes']


def build_template(path: str) -> str:
    document = docx.Document()
    body = document.element.body
    for xml in [CLIENT_NAME, TOTAL, PAYMENT_METHOD, ORDER_ITEMS, RETURN_CHANGE, NOTES]:
        body.insert(len(body) - 1, parse_xml(xml))
    document.save(path)
    return path
