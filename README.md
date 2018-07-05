## ExcelUtil 

Add these snippets in your `pom.xml`

```xml
...
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>
...
<dependency>
    <groupId>com.github.CriptoCosmo</groupId>
    <artifactId>excel-util</artifactId>
    <version>master</version>
</dependency>
...
```



------

### Example - Excel File 

![Example Excel File](.//images//excel.png)

### Example - ModelEntity

```java
@ExcelEntity
public class ExcelRow {
	
	private static final String PATTERN = "dd MMMM yyyy";
	private static final String LOCALE = "it";
	
	@ExcelField private String nome ;
	@ExcelField private String cognome ;
	@ExcelField private Double classe ;
	
	// CUSTOM PATTERN AND LOCALE 
	@ExcelField(value=PATTERN,locale=LOCALE) 
	private Date data ;
    
	// DEFAULT DATE 
	@ExcelField 
	private Date data2 ;

}
```

### Example - BasicUsage		

```java
public static void main(String[] args) throws Exception {
	
    String path = "C:\Users\Cosmo\Desktop\sample.xlsx" ;
	
	ExcelReader<ExcelRow> excelReader = new ExcelReaderImpl<ExcelRow>(ExcelRow.class,path);
	
	for (ExcelRow excelRow : excelReader.readRows()) {
		System.out.println(excelRow);
	}
}
```
### Result 

```log
ExcelRow [nome=Mario, cognome=Arbola, classe=Quinta, data=19 Luglio 1996] 
```

