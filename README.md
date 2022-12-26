# excelUtil

grandle

allprojects {
	repositories {
		maven { url 'https://jitpack.io' }
	}
}

dependencies {
	implementation 'com.github.hanmailco34:excelUtil:0.0.2'
}

maven

<repositories>
	<repository>
	    <id>jitpack.io</id>
	    <url>https://jitpack.io</url>
	</repository>
</repositories>

<dependency>
    <groupId>com.github.hanmailco34</groupId>
    <artifactId>excelUtil</artifactId>
    <version>0.0.2</version>
</dependency>

예제
엑셀로 만들 VO 데이터에-------------1

  @ExcelColumn(headerName = "아이디")
  
	private String userId;
	
서비스 단에서

-----------------------2

ExcelUtil<만든vo> excelFile = new ExcelUtil(자료리스트, 만든vo.class);

----------------------3

excelFile.write(HttpServletResponse response);를 넣으면 프론트에서 사용

excelFile.write(String filpaht);를 넣으면 백단에서 다운로드
