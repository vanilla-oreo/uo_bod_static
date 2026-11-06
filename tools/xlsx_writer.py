"""Pure-standard-library OOXML spreadsheet writer for UO BOD management.
Writes five sheets with formula-driven redemption matrix. No pip installation needed.
"""
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
from pathlib import Path

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'

def col(num):
    chars=''
    while num:
        num, r=divmod(num-1,26);chars=chr(65+r)+chars
    return chars

def cell(value,address,style=0):
    if value is None:return ''
    st=f' s="{style}"' if style else ''
    if isinstance(value,str) and value.startswith('='):
        return f'<c r="{address}"{st}><f>{escape(value[1:])}</f></c>'
    if isinstance(value,(int,float)) and not isinstance(value,bool):
        return f'<c r="{address}"{st}><v>{value}</v></c>'
    txt=escape(str(value))
    return f'<c r="{address}"{st} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'

def sheet(rows,freeze_cols=0,custom_widths=None):
    max_cols=max((len(r) for r in rows),default=1)
    widths=custom_widths or {}
    cols=''.join(f'<col min="{i}" max="{i}" width="{widths.get(i,22)}" customWidth="1"/>' for i in range(1,max_cols+1))
    frozen=(f'<sheetViews><sheetView workbookViewId="0"><pane xSplit="{freeze_cols}" ySplit="1" topLeftCell="{col(freeze_cols+1)}2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' if freeze_cols else '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    lines=[]
    for i,row in enumerate(rows,1):
        cells=''.join(cell(v,f'{col(j)}{i}',1 if i==1 else 0) for j,v in enumerate(row,1))
        lines.append(f'<row r="{i}"{ " ht=\"30\" customHeight=\"1\"" if i==1 else "" }>{cells}</row>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{MAIN}">{frozen}<cols>{cols}</cols><sheetData>{"".join(lines)}</sheetData>'
            f'<autoFilter ref="A1:{col(max_cols)}{max(len(rows),2)}"/></worksheet>').encode('utf-8')

def write(path,sheets):
    path=Path(path)
    content_types=['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>','<Default Extension="xml" ContentType="application/xml"/>','<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>','<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    content_types.extend(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(sheets)+1))
    types='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+''.join(content_types)+'</Types>'
    book='<workbook xmlns="'+MAIN+'" xmlns:r="'+REL+'"><sheets>'+''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i,(name,*_) in enumerate(sheets,1))+'</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
    rls='<Relationships xmlns="'+PKG+'">'+''.join(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))+f'<Relationship Id="rId{len(sheets)+1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'
    styles=f'''<styleSheet xmlns="{MAIN}"><fonts count="2"><font><sz val="11"/><name val="Microsoft JhengHei"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft JhengHei"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF183153"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>'''
    rootrels=f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',types)
        z.writestr('_rels/.rels',rootrels)
        z.writestr('xl/workbook.xml',book)
        z.writestr('xl/_rels/workbook.xml.rels',rls)
        z.writestr('xl/styles.xml',styles)
        for i,(_name,rows,freeze,widths) in enumerate(sheets,1):
            z.writestr(f'xl/worksheets/sheet{i}.xml',sheet(rows,freeze,widths))
