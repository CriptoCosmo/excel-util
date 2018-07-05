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

```shell
ExcelRow [nome=Mario, cognome=Arbola, classe=5.0, data=Fri Jul 19 00:00:00 CEST 1996, data2=Fri Jan 19 00:00:00 CET 1900]

ExcelRow [nome=Gabriele, cognome=Cipolloni, classe=5.0, data=Sat Nov 11 00:00:00 CET 1995, data2=Fri Jan 19 00:00:00 CET 1900]
```
