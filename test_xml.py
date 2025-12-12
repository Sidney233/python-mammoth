import results
import transforms
from office_xml import read, read_str
import lists
from xmlparser import XmlElement
import documents
from styles_xml import Styles

class _ReadResult(object):
    @staticmethod
    def concat(results):
        return _ReadResult(
            lists.flat_map(lambda result: result.elements, results),
            lists.flat_map(lambda result: result.extra, results),
            lists.flat_map(lambda result: result.messages, results))

    @staticmethod
    def map_results(first, second, func):
        return _ReadResult(
            [func(first.elements, second.elements)],
            first.extra + second.extra,
            first.messages + second.messages)

    def __init__(self, elements, extra, messages):
        self.elements = elements
        self.extra = extra
        self.messages = messages

    def map(self, func):
        elements = func(self.elements)
        if not isinstance(elements, list):
            elements = [elements]
        return _ReadResult(
            elements,
            self.extra,
            self.messages)

    def flat_map(self, func):
        result = func(self.elements)
        return _ReadResult(
            result.elements,
            self.extra + result.extra,
            self.messages + result.messages)

    def to_extra(self):
        return _ReadResult([], _concat(self.extra, self.elements), self.messages)

    def append_extra(self):
        return _ReadResult(_concat(self.elements, self.extra), [], self.messages)


def _concat(*values):
    result = []
    for value in values:
        for element in value:
            result.append(element)
    return result


def _read_xml_elements(nodes):
    elements = filter(lambda node: isinstance(node, XmlElement), nodes)
    return _ReadResult.concat(lists.map(read, elements))


def calculate_row_spans(rows):
    unexpected_non_rows = any(
        not isinstance(row, documents.TableRow)
        for row in rows
    )
    if unexpected_non_rows:
        rows = remove_unmerged_table_cells(rows)
        return _elements_result_with_messages(rows, [results.warning(
            "unexpected non-row element in table, cell merging may be incorrect"
        )])

    unexpected_non_cells = any(
        not isinstance(cell, documents.TableCellUnmerged)
        for row in rows
        for cell in row.children
    )
    if unexpected_non_cells:
        rows = remove_unmerged_table_cells(rows)
        return _elements_result_with_messages(rows, [results.warning(
            "unexpected non-cell element in table row, cell merging may be incorrect"
        )])

    columns = {}
    for row in rows:
        cell_index = 0
        for cell in row.children:
            if cell.vmerge and cell_index in columns:
                columns[cell_index].rowspan += 1
            else:
                columns[cell_index] = cell
                cell.vmerge = False
            cell_index += cell.colspan

    for row in rows:
        row.children = [
            documents.table_cell(
                children=cell.children,
                colspan=cell.colspan,
                rowspan=cell.rowspan,
            )
            for cell in row.children
            if not cell.vmerge
        ]

    return _success(rows)


def _success(elements):
    if not isinstance(elements, list):
        elements = [elements]
    return _ReadResult(elements, [], [])


def remove_unmerged_table_cells(rows):
    return list(map(
        transforms.element_of_type(
            documents.TableCellUnmerged,
            lambda cell: documents.table_cell(
                children=cell.children,
                colspan=cell.colspan,
                rowspan=cell.rowspan,
            ),
        ),
        rows,
    ))


def _elements_result_with_messages(elements, messages):
    return _ReadResult(elements, [], messages)


table = read_str(
    """<w:tbl xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"><w:tblPr><w:tblStyle w:val="GridTable1Light-Accent1"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3116"/><w:gridCol w:w="3117"/><w:gridCol w:w="3117"/></w:tblGrid><w:tr w:rsidR="00D43D84" w14:paraId="1DF9F8A5" w14:textId="77777777" w:rsidTr="00D43D84"><w:trPr><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:tcW w:w="3116" w:type="dxa"/></w:tcPr><w:p w14:paraId="3A047453" w14:textId="1F41A2BF" w:rsidR="00D43D84" w:rsidRDefault="00D43D84"><w:r><w:t>Header 0.0</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="57EF64D9" w14:textId="295B1EF6" w:rsidR="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Header 0.1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="15FCCABA" w14:textId="61A65E49" w:rsidR="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Header 0.2</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00D43D84" w14:paraId="30A59335" w14:textId="77777777" w:rsidTr="00D43D84"><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:tcW w:w="3116" w:type="dxa"/></w:tcPr><w:p w14:paraId="19C498D4" w14:textId="02BB9B49" w:rsidR="00D43D84" w:rsidRPr="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr><w:t>Cell 1.0</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="2627AA0D" w14:textId="19DB13B9" w:rsidR="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Cell 1.1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="5AB36E91" w14:textId="09925B27" w:rsidR="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Cell 1.2</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00D43D84" w:rsidRPr="00D43D84" w14:paraId="65971744" w14:textId="77777777" w:rsidTr="00D43D84"><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:tcW w:w="3116" w:type="dxa"/></w:tcPr><w:p w14:paraId="5B0D2C9F" w14:textId="4100C0D9" w:rsidR="00D43D84" w:rsidRPr="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr><w:t>Cell 2.0</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="2375BDD5" w14:textId="16CBB136" w:rsidR="00D43D84" w:rsidRPr="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Cell 2.1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3117" w:type="dxa"/></w:tcPr><w:p w14:paraId="770067B2" w14:textId="34A97FFE" w:rsidR="00D43D84" w:rsidRPr="00D43D84" w:rsidRDefault="00D43D84"><w:pPr><w:cnfStyle w:val="000000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/></w:pPr><w:r><w:t>Cell 2.2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>""")
res = _ReadResult.map_results(
    Styles.create(),
    _read_xml_elements(table.children)
        .flat_map(calculate_row_spans),

    lambda style, children: documents.table(
        children=children,
        style_id=style[0],
        style_name=style[1],
    ),
)

print(res)