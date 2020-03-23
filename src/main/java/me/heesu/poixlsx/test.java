/*
작업 이루어지는 프로세스 별 정리
*/

/*1.  xlsx 지원 엑셀 생성  SAX event api 이용 (해당 메서드 내에서 parser를 통해서 StringVal로 대용량 excel read => List<String[]>
  18만건, 10Mb기준 10초 이내
)*/
public void createXlsxExcelFileBySax(String filePath, HttpServletResponse response) throws Exception {
  /* batch에서 추출된 데이터 filePath로 엑셀템플릿파일 지정 */
  String pathPrefixForTest = "C:\\";
  File file = new File(pathPrefixForTest + filePath);
  DateFormat sdf = new SimpleDateFormat("_yyyyMMddhhmmss");
    Date toDay = new Date();
    String tempDate = sdf.format(toDay);
    String fileName = "PromotionData" + tempDate;

  if (!file.exists()) {
    throw new BizMsgException("파일이 존재하지 않습니다.");
  }

      OPCPackage pkg = OPCPackage.open(file);
      XSSFReader r = new XSSFReader( pkg );
      SharedStringsTable sst = r.getSharedStringsTable();
      XMLReader parser = fetchSheetParser(sst);

      //첫번째 시트에 대한 처리 (데이터추출요청별 시트 하나만 생성)
      InputStream sheet = r.getSheetsData().next();
      InputSource sheetSource = new InputSource(sheet);
      parser.parse(sheetSource);
      sheet.close();

      List<String[]> res = new LinkedList<String[]>();
    res = SheetHandler.getRowCache();

    createXlsxExcelFileTest(res, response);

  }

  //eventapi로 읽어온 데이터에 대해서 엑셀파일 생성
	public void createXlsxExcelFileTest (List<String[]> resultList, HttpServletResponse response) {
		try {
			String fileName = "OOM_TEST";
			SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(100); // 메모리에 들고있는 최대 row수 설정
			Sheet sheet = sxssfWorkbook.createSheet(); // Sheet
	        Row row = null; // Row
	        Cell cell = null; // Cell

			if (resultList != null && !resultList.isEmpty()) {
	            row = sheet.createRow(0);
	            String[] headerResult = resultList.get(0);
	            int totalRowNum = (resultList != null) ? resultList.size() : 0;
	            int totalColNum = (headerResult != null) ? headerResult.length : 0 ;
	            // 첫 row - 컬럼명
	            for (int i = 0;i < totalColNum ; i++) {
	                cell = row.createCell(i);
	                cell.setCellValue(headerResult[i]);
	            }

	            // 그 이후 row
	            for (int i = 0; i < totalRowNum; i++) {
	                row = sheet.createRow(i+1);
	                String[] result = resultList.get(i);
	                for (int j = 0; j< totalColNum; j++) {
	                    String cellVal = result[j];
	                    cell = row.createCell(j);
	                    if (cellVal != null) {
	                    	cell.setCellValue(cellVal);
	                    } else {
	                        cell.setCellValue("");
	                    }
	                }
	                System.out.println(i);
	            }
	        }

			/* 엑셀파일 생성 */
	        response.setHeader("Content-disposition", "attachment;filename=" + fileName + XLSX_EXTENTION_NAME);
	        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

	        sxssfWorkbook.write(response.getOutputStream());
	        ((SXSSFWorkbook)sxssfWorkbook).dispose();

		}catch(Exception e) {
			//: TODO
			e.printStackTrace();
			throw new BizMsgException("엑셀파일 생성 오류.");
		}
	}

	/* parser 세팅 */
	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        //XMLReader parser = SAXHelper.newXMLReader();
        SAXParserFactory parserFactory = SAXParserFactory.newInstance();
        SAXParser parser = parserFactory.newSAXParser();
        XMLReader xmlParser = parser.getXMLReader();

		ContentHandler handler = new SheetHandler(sst);
		xmlParser.setContentHandler(handler);
        return xmlParser;
    }


	
	/* XSSF로 엑셀 read시 out of memory 발생하는 것 방지하기 위해서 따로 poi에서 제공하는 sax api를 이용해서 read 직접 처리  */
	private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean inlineStr;

        private static final String ROW_EVENT = "row";
        private static final String CELL_EVENT = "c";

        private static List<String> cellCache = new LinkedList<String>();
        private static List<String[]> rowCache = new LinkedList<String[]>();

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if(CELL_EVENT.equals(name)) {
                // Print the cell reference
                //System.out.print(attributes.getValue("r") + " - ");
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");

                nextIsString = (cellType != null && cellType.equals("s"));
                inlineStr = (cellType != null && cellType.equals("inlineStr"));

            }else if(ROW_EVENT.equals(name)) {
            	if(!cellCache.isEmpty()) {
            		rowCache.add(cellCache.toArray(new String[cellCache.size()]));
            	}
            	cellCache.clear();
            }

            // Clear contents cache
            lastContents = "";
        }
        @Override
        public void endElement(String uri, String localName, String name)
                throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();  // sst.getItemAt(idx).getString();
                nextIsString = false;
            }
            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v") || (inlineStr && name.equals("c"))) {
                //System.out.println(lastContents);

                /* cell data 저장 */
                cellCache.add(lastContents);

            }
        }
        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }

        @Override
        public void endDocument() throws SAXException {
        	/* doc parsing 완료 후, sxssfworkbook 생성..? */

        	//System.out.println(res.toString());
        	System.out.println("END DOCU######################");
        }

        public static List<String[]> getRowCache() {
        	return rowCache;
        }
    }
