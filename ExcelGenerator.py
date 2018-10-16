#!/bin/env python
# -*- coding: utf-8-sig -*-

import xml.etree.ElementTree as ET
import sys, os, zipfile, datetime, platform

log = open('error.log', 'w')

zip_ext = '.zip'
xml_ext = '.xml'
sig_ext = '.sig'

DIR = os.getcwd()
valid_files = []
list_first_arch = []
list_second_arch = []


HEAD = """<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Serj</Author>
  <LastAuthor>Windows User</LastAuthor>
  <LastPrinted>2018-10-10T09:29:15Z</LastPrinted>
  <Created>2018-10-10T04:27:37Z</Created>
  <LastSaved>2018-10-14T04:32:55Z</LastSaved>
  <Version>15.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>11145</WindowHeight>
  <WindowWidth>28800</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>465</WindowTopY>
  <TabRatio>500</TabRatio>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
  <DisplayInkNotes>False</DisplayInkNotes>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s42" ss:Name="Обычный 2">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Arial Cyr" x:CharSet="204"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s64" ss:Parent="s42">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Size="14"/>
   <Interior/>
  </Style>
  <Style ss:ID="s65" ss:Parent="s42">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"/>
   <Interior/>
  </Style>
  <Style ss:ID="s66" ss:Parent="s42">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Size="14"/>
   <Interior/>
   <NumberFormat ss:Format="#&quot; &quot;?/?"/>
  </Style>
  <Style ss:ID="s67" ss:Parent="s42">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"/>
   <Interior ss:Color="#92D050" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s68">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s69">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s70">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#000000"/>
   <NumberFormat ss:Format="#&quot; &quot;?/?"/>
  </Style>
  <Style ss:ID="s71">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s72">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="Fixed"/>
  </Style>
  <Style ss:ID="s73">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="0"/>
  </Style>
  <Style ss:ID="s74">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Color="#000000"/>
   <Interior/>
  </Style>
  <Style ss:ID="s75">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="Fixed"/>
  </Style>
  <Style ss:ID="s76">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:CharSet="204" x:Family="Roman"
    ss:Color="#000000"/>
   <Interior ss:Color="#92D050" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="Fixed"/>
  </Style>
  <Style ss:ID="s79">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <NumberFormat ss:Format="Fixed"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Лист2">
  <Table ss:ExpandedColumnCount="11" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="66" ss:DefaultRowHeight="15.75">
   <Column ss:AutoFitWidth="0" ss:Width="30.75"/>
   <Column ss:Width="234"/>
   <Column ss:Width="413.25"/>
   <Column ss:Width="36.75"/>
   <Column ss:Width="30.75"/>
   <Column ss:Width="47.25"/>
   <Column ss:Index="8" ss:AutoFitWidth="0" ss:Width="48.75"/>
   <Column ss:Width="60.75"/>
   <Column ss:Width="57" ss:Span="1"/>
   <Row ss:Height="63.75">
    <Cell ss:StyleID="s64"><Data ss:Type="String">№ </Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">Собственник помещения </Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String"> Документ, подтверждающий право собственности на жилое помещение </Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Доля</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">S</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String"> </Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">Доля в праве общей собственности на общее имущество, % </Data></Cell>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"><Data ss:Type="String">% от общего числа голосов</Data></Cell>
    <Cell ss:StyleID="s67"><Data ss:Type="String">Бюллетень (ЗА +, ПРОТИВ -) </Data></Cell>
    <Cell ss:StyleID="s67"><Data ss:Type="String">Бюллетень (ЗА +, ПРОТИВ -) </Data></Cell>
   </Row>
"""
BODY = """   <Row>
    <Cell ss:StyleID="s68"><Data ss:Type="String">{apartment}</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">{fio}</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">{doc}</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="Number">{share}</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="Number">{area}</Data></Cell>
    <Cell ss:StyleID="s72" ss:Formula="=RC[-1]*RC[-2]"><Data ss:Type="Number">0</Data></Cell>
    <Cell ss:StyleID="s73"/>
    <Cell ss:StyleID="s74"/>
    <Cell ss:StyleID="s75"/>
    <Cell ss:StyleID="s76"/>
    <Cell ss:StyleID="s76"/>
   </Row>
"""

FOOTER = """   <Row>
    <Cell ss:Index="6" ss:StyleID="s79"><Data ss:Type="Number">1021</Data></Cell>
    <Cell ss:StyleID="s73" ss:Formula="=RC[-1]*1000"><Data ss:Type="Number">1021000</Data></Cell>
    <Cell ss:StyleID="s71"/>
    <Cell ss:StyleID="s75" ss:Formula="=SUM(R[-1]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>
    <Cell ss:StyleID="s76" ss:Formula="=IF(RC[-2]=&quot;+&quot;,RC[-1],0)"><Data
      ss:Type="Number">0</Data></Cell>
    <Cell ss:StyleID="s76" ss:Formula="=IF(RC[-3]=&quot;+&quot;,RC[-4],0)"><Data
      ss:Type="Number">0</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-4</HorizontalResolution>
    <VerticalResolution>-4</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>1</ActiveRow>
     <ActiveCol>1</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
"""

table = []

def get_child_by_name(node, name):
    for n in node:
        if n.tag.split('}')[1] == name:
            yield n


print('='*80)
print('Find files for convertation')
print('='*80)

for i in range(2):
    for item in os.listdir(DIR): # loop through items in dir
        if item.endswith(zip_ext): # check for ".zip" extension
            file_name = os.path.abspath(item) # get full path of files
            zip_ref = zipfile.ZipFile(file_name) # create zipfile object
            zip_ref.extractall('./') # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) # delete zipped file

for f in os.listdir(DIR):
    if f.lower().endswith(xml_ext):
        valid_files.append(f)
    elif f.lower().endswith(sig_ext):
        os.remove(f)

table = {}
for x in valid_files:
    print(x)
    try:
        tree = ET.parse(x)
        root = tree.getroot()
        realty = root.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Realty")
        reestr = root.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}ReestrExtract")
        flat = realty.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Flat")
        area = float(flat.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Area").text)
        address = flat.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Address")
        apartment = address.find("{urn://x-artefacts-rosreestr-ru/commons/complex-types/address-output/4.0.1}Apartment").attrib['Value']

        #Собственники
        rights = reestr.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}ExtractObjectRight")
        rights = rights.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}ExtractObject")
        rights = rights.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}ObjectRight")
        table[int(apartment)] = []
        for r in rights.findall("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Right"):
            registration = r.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Registration")
            doc_name = registration.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Name").text#.encode('utf-8')
            share = registration.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}ShareText")
            share = share.text if share is not None else '1'
            owner = r.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Owner")
            person = owner.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Person")
            fio = person.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}FIO")
            fn = fio.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}First").text#.encode('utf-8')
            ln = fio.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Surname").text#.encode('utf-8')
            p = fio.find("{urn://x-artefacts-rosreestr-ru/outgoing/kpoks/4.0.1}Patronymic").text#.encode('utf-8')
            if len(table[int(apartment)]) > 0: 
                ap_num = ''
            else:
                ap_num = apartment
            table[int(apartment)].append({'apartment': ap_num, 
                                         'fio': "{} {} {}".format(ln, fn, p), 
                                         'doc': doc_name, 'share': eval(share), 
                                         'area': area})
    except Exception as e:
        log.write("{} is invalid. Error: {}\n".format(x, e))

table_string = ''
for k in sorted(table):
    for j in table[k]:
        table_string += BODY.format(**j)

OUT = HEAD + table_string + FOOTER

out_file_name = "{}.xls".format(datetime.datetime.now().strftime('%d%m%Y%H%M%S'))
outf = open(out_file_name, 'w', encoding='utf-8-sig')
for i in OUT:
    outf.write(i)
outf.close()
log.close()

input("Press Enter to continue...")

if platform.system() == 'Windows':
    os.system('start excel.exe {}'.format(out_file_name))
else:
    os.system("open -a'Microsoft Excel.app' '{}'".format(out_file_name))
