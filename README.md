# excelUtil

예제
엑셀로 만들 VO 데이터에 
  @ExcelColumn(headerName = "아이디")
	private String userId;
서비스 단에서
ExcelUtil<만든vo> excelFile = new ExcelUtil(자료리스트, 만든vo.class);

excelFile.write(HttpServletResponse response);를 넣으면 프론트에서 사용
excelFile.write(String filpaht);를 넣으면 백단에서 다운로드
