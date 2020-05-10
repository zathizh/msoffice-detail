# msoffice-detail
Newesst Microsoft Office Documents (DOCX, XLSX, PPTX) are built as an archive file. It can be extracted into collection of XML files.
core.xml on docProps contains the MS Document properties like (Title, Subject, Creator, Keywords, Description, etc.)

Below sample structure was used to extract the FILE properties from the MS Document.

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
	xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
	xmlns:dc="http://purl.org/dc/elements/1.1/"
	xmlns:dcterms="http://purl.org/dc/terms/" 
	xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title></dc:title>
<dc:subject></dc:subject>
<dc:creator>Creator/dc:creator>
<cp:keywords>Keywords</cp:keywords>
<dc:description>Description</dc:description>
<cp:lastModifiedBy>Last Modified By</cp:lastModifiedBy>
<cp:revision>Revision</cp:revision>
<dcterms:created xsi:type="dcterms:W3CDTF">Created Timestamp</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">Created Timestamp</dcterms:modified></cp:coreProperties

msoffice-det.py can be use to extract these details from a Single MS Document or Folder/Directory which contains Multiple MS Documents 
