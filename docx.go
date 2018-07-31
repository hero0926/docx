package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"io/ioutil"
	"os"
	"strings"
)

//Contains functions to work with data from a zip file
type ZipData interface {
	files() []*zip.File
	close() error
}

//Type for in memory zip files
type ZipInMemory struct {
	data *zip.Reader
}

func (d ZipInMemory) files() []*zip.File {
	return d.data.File
}

//Since there is nothing to close for in memory, just nil the data and return nil
func (d ZipInMemory) close() error {
	d.data = nil
	return nil
}

//Type for zip files read from disk
type ZipFile struct {
	data *zip.ReadCloser
}

func (d ZipFile) files() []*zip.File {
	return d.data.File
}

func (d ZipFile) close() error {
	return d.data.Close()
}

type ReplaceDocx struct {
	zipReader ZipData
	content   string
	links     string
	headers   map[string]string
	footers   map[string]string
}

func (r *ReplaceDocx) Editable() *Docx {
	return &Docx{
		files:   r.zipReader.files(),
		content: r.content,
		links:   r.links,
		headers: r.headers,
		footers: r.footers,
	}
}

func (r *ReplaceDocx) Close() error {
	return r.zipReader.close()
}

type Docx struct {
	files   []*zip.File
	content string
	links   string
	headers map[string]string
	footers map[string]string
}

func (d *Docx) GetText() (text string) {
	/* 이거 이런식으로 리턴됨 <>를 제거해보자.
	   <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se wp14"><w:body><w:p w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidRDefault="00067713" w:rsidP="00067713"><w:pPr><w:widowControl/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:before="450" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t>계정</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t>발급</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t>및</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:spacing w:val="-2"/><w:kern w:val="0"/><w:sz w:val="35"/><w:szCs w:val="35"/></w:rPr><w:t>관리</w:t></w:r></w:p><w:p w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidRDefault="00067713" w:rsidP="00067713"><w:pPr><w:widowControl/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:before="150" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr><w:t xml:space="preserve">Long term user </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr><w:t>계정</w:t></w:r></w:p><w:p w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidRDefault="00067713" w:rsidP="00067713"><w:pPr><w:widowControl/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:before="150" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>UI</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>개발팀은</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>이미</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> Long term user </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>공용</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>계정이</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>발급되어</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>있습니다</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="444444"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>.</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="13227" w:type="dxa"/><w:tblCellSpacing w:w="0" w:type="dxa"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="13227"/></w:tblGrid><w:tr w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidTr="008F23F0"><w:trPr><w:tblCellSpacing w:w="0" w:type="dxa"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="13002" w:type="dxa"/><w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders><w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="225" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar><w:vAlign w:val="bottom"/><w:hideMark/></w:tcPr><w:p w:rsidR="00067713" w:rsidRDefault="00067713" w:rsidP="008F23F0"><w:pPr><w:widowControl/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:after="0" w:line="300" w:lineRule="atLeast"/><w:jc w:val="left"/><w:textAlignment w:val="baseline"/><w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="굴림체" w:hAnsi="Consolas" w:cs="Consolas"/><w:kern w:val="0"/><w:szCs w:val="20"/><w:bdr w:val="none" w:sz="0" w:space="0" w:color="auto" w:frame="1"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="굴림체" w:hAnsi="Consolas" w:cs="Consolas"/><w:kern w:val="0"/><w:szCs w:val="20"/><w:bdr w:val="none" w:sz="0" w:space="0" w:color="auto" w:frame="1"/></w:rPr><w:t>U2FsdGVkX1+rsu02MYxnNF49h0XCN8NPeZbVM2RJPRklCnQjLZishDK2j3fXFvr5</w:t></w:r></w:p><w:p w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidRDefault="00067713" w:rsidP="008F23F0"><w:pPr><w:widowControl/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:after="0" w:line="300" w:lineRule="atLeast"/><w:jc w:val="left"/><w:textAlignment w:val="baseline"/><w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="굴림" w:hAnsi="Consolas" w:cs="Consolas"/><w:kern w:val="0"/><w:szCs w:val="20"/></w:rPr></w:pPr></w:p></w:tc></w:tr></w:tbl><w:p w:rsidR="00067713" w:rsidRPr="00B11D25" w:rsidRDefault="00067713" w:rsidP="00067713"><w:pPr><w:widowControl/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:before="450" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr><w:t xml:space="preserve">Temporary user </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr><w:t>계정</w:t></w:r></w:p><w:p w:rsidR="00067713" w:rsidRDefault="00067713" w:rsidP="00067713"><w:pPr><w:widowControl/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:spacing w:before="150" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">Long term user </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>계정은</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:proofErr w:type="spellStart"/><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>필요할때</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>아래의</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>절차대로</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>신청합니다</w:t></w:r><w:r w:rsidRPr="00B11D25"><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="222222"/><w:kern w:val="0"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>.</w:t></w:r></w:p><w:p w:rsidR="00D70EA5" w:rsidRDefault="00D70EA5"><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:sectPr w:rsidR="00D70EA5"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>
	*/

	return d.content
}

func (d *Docx) ReplaceRaw(oldString string, newString string, num int) {
	d.content = strings.Replace(d.content, oldString, newString, num)
}

func (d *Docx) Replace(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.content = strings.Replace(d.content, oldString, newString, num)

	return nil
}

func (d *Docx) ReplaceLink(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.links = strings.Replace(d.links, oldString, newString, num)

	return nil
}

func (d *Docx) ReplaceHeader(oldString string, newString string) (err error) {
	return replaceHeaderFooter(d.headers, oldString, newString)
}

func (d *Docx) ReplaceFooter(oldString string, newString string) (err error) {
	return replaceHeaderFooter(d.footers, oldString, newString)
}

func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

func (d *Docx) Write(ioWriter io.Writer) (err error) {
	w := zip.NewWriter(ioWriter)
	for _, file := range d.files {
		var writer io.Writer
		var readCloser io.ReadCloser

		writer, err = w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		if file.Name == "word/document.xml" {
			writer.Write([]byte(d.content))
		} else if file.Name == "word/_rels/document.xml.rels" {
			writer.Write([]byte(d.links))
		} else if strings.Contains(file.Name, "header") && d.headers[file.Name] != "" {
			writer.Write([]byte(d.headers[file.Name]))
		} else if strings.Contains(file.Name, "footer") && d.footers[file.Name] != "" {
			writer.Write([]byte(d.footers[file.Name]))
		} else {
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return
}

func replaceHeaderFooter(headerFooter map[string]string, oldString string, newString string) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}

	for k := range headerFooter {
		headerFooter[k] = strings.Replace(headerFooter[k], oldString, newString, -1)
	}

	return nil
}

func ReadDocxFromMemory(data io.ReaderAt, size int64) (*ReplaceDocx, error) {
	reader, err := zip.NewReader(data, size)
	if err != nil {
		return nil, err
	}
	zipData := ZipInMemory{data: reader}
	return ReadDocx(zipData)
}

func ReadDocxFile(path string) (*ReplaceDocx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	zipData := ZipFile{data: reader}
	return ReadDocx(zipData)
}

func ReadDocx(reader ZipData) (*ReplaceDocx, error) {
	content, err := readText(reader.files())
	if err != nil {
		return nil, err
	}

	links, err := readLinks(reader.files())
	if err != nil {
		return nil, err
	}

	headers, footers, _ := readHeaderFooter(reader.files())
	return &ReplaceDocx{zipReader: reader, content: content, links: links, headers: headers, footers: footers}, nil
}

func readHeaderFooter(files []*zip.File) (headerText map[string]string, footerText map[string]string, err error) {

	h, f, err := retrieveHeaderFooterDoc(files)

	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	headerText, err = buildHeaderFooter(h)
	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	footerText, err = buildHeaderFooter(f)
	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	return headerText, footerText, err
}

func buildHeaderFooter(headerFooter []*zip.File) (map[string]string, error) {

	headerFooterText := make(map[string]string)
	for _, element := range headerFooter {
		documentReader, err := element.Open()
		if err != nil {
			return map[string]string{}, err
		}

		text, err := wordDocToString(documentReader)
		if err != nil {
			return map[string]string{}, err
		}

		headerFooterText[element.Name] = text
	}

	return headerFooterText, nil
}

func readText(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveWordDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func readLinks(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveLinkDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func wordDocToString(reader io.Reader) (string, error) {
	b, err := ioutil.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

func retrieveWordDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml file not found")
	}
	return
}

func retrieveLinkDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/_rels/document.xml.rels" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml.rels file not found")
	}
	return
}

func retrieveHeaderFooterDoc(files []*zip.File) (headers []*zip.File, footers []*zip.File, err error) {
	for _, f := range files {

		if strings.Contains(f.Name, "header") {
			headers = append(headers, f)
		}
		if strings.Contains(f.Name, "footer") {
			footers = append(footers, f)
		}
	}
	if len(headers) == 0 && len(footers) == 0 {
		err = errors.New("headers[1-3].xml file not found and footers[1-3].xml file not found.")
	}
	return
}

func streamToByte(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	buf.ReadFrom(stream)
	return buf.Bytes()
}

func encode(s string) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(bufio.NewWriter(&b))
	if err := enc.Encode(s); err != nil {
		return s, err
	}
	output := strings.Replace(b.String(), "<string>", "", 1) // remove string tag
	output = strings.Replace(output, "</string>", "", 1)
	output = strings.Replace(output, "&#xD;&#xA;", "<w:br/>", -1) // \r\n => newline
	return output, nil
}
