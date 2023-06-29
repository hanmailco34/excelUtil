# ExcelUtil 라이브러리

ExcelUtil은 자바에서 엑셀 파일을 처리하기 위한 유틸리티 라이브러리입니다. 이 라이브러리를 사용하면 리스트를 엑셀 파일로 변환하거나, 엑셀 파일을 리스트로 변환하는 등 다양한 기능을 수행할 수 있습니다.

## 설치

### Gradle

```groovy
allprojects {
    repositories {
        maven { url 'https://jitpack.io' }
    }
}

dependencies {
    implementation 'com.github.hanmailco34:excelUtil:0.0.5'
}
```

### Maven

```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>

<dependencies>
    <dependency>
        <groupId>com.github.hanmailco34</groupId>
        <artifactId>excelUtil</artifactId>
        <version>0.0.5</version>
    </dependency>
</dependencies>
```

## 사용 예제

### 1. 리스트에서 엑셀로 변환하기

```java
// TestVO 클래스 생성
public class TestVO {
    @ExcelColumn(headerName = "아이디")
    private String userId;

    // Getter, Setter 및 기타 필요한 메서드
}

// 자료 리스트 생성
List<TestVO> dataList = new ArrayList<>();
TestVO data1 = new TestVO();
data1.setUserId("user1");
dataList.add(data1);
TestVO data2 = new TestVO();
data2.setUserId("user2");
dataList.add(data2);

// ExcelUtil 인스턴스 생성
ExcelUtil<TestVO> excelFile = new ExcelUtil<>(dataList, TestVO.class);

// 엑셀 파일 생성 및 다운로드
excelFile.write("파일경로/파일명.xlsx");
```

### 2. 엑셀에서 리스트로 변환하기

```java
// 엑셀 파일 읽어오기
File excelFile = new File("파일경로/파일명.xlsx");

// ExcelUtil 인스턴스 생성
ExcelUtil<TestVO> excelUtil = new ExcelUtil<>(excelFile, TestVO.class);

// 리스트로 변환
List<TestVO> dataList = excelUtil.convertToList();

// 변환된 데이터 활용
for (TestVO data : dataList) {
    System.out.println("아이디: " + data.getUserId());
}
```

### 3. 여러 개의 리스트로 하나의 엑셀 파일에서 여러 시트 만들기

```java
// TestVO1 클래스 생성
public class TestVO1 {
    @ExcelColumn(headerName = "이름")
    private String name;

    // Getter, Setter 및 기타 필요한 메서드
}

// TestVO2 클래스 생성
public class TestVO2 {
    @ExcelColumn(headerName = "학번")
    private int studentId;

    // Getter, Setter 및 기타 필요한 메서드
}

// 자료 리스트1 생성
List<TestVO1> dataList1 = new ArrayList<>();
TestVO1 data1 = new TestVO1();
data1.setName("홍길동");
dataList1.add(data1);
TestVO1 data2 = new TestVO1();
data2.setName("이순신");
dataList1

.add(data2);

// 자료 리스트2 생성
List<TestVO2> dataList2 = new ArrayList<>();
TestVO2 data3 = new TestVO2();
data3.setStudentId(20210001);
dataList2.add(data3);
TestVO2 data4 = new TestVO2();
data4.setStudentId(20210002);
dataList2.add(data4);

// ExcelUtil 인스턴스 생성
ExcelUtil<TestVO1> excelFile1 = new ExcelUtil<>(dataList1, TestVO1.class);
ExcelUtil<TestVO2> excelFile2 = new ExcelUtil<>(dataList2, TestVO2.class);

// 리스트에 순서대로 추가
List<ExcelUtil<?>> excelList = new ArrayList<>();
excelList.add(excelFile1);
excelList.add(excelFile2);

// 엑셀 파일 생성 및 다운로드
excelUtil.write("파일경로/파일명.xlsx");
```

위의 예제 코드는 ExcelUtil 라이브러리의 주요 기능을 설명하기 위한 것입니다.
실제 사용에는 필요에 따라 코드를 수정하고, 데이터를 적절하게 처리해야 합니다.
