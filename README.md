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

### Example entity to map excel file

```java
@ExcelEntity
public class ExcelRow {

	@ExcelField private String nome ;
	@ExcelField private String cognome ;
	@ExcelField private String classe ;
	@ExcelField private String data ;
	
    // NOT NEED
	@Override
	public String toString() {
        return "ExcelRow [nome=" + nome + ", cognome=" + cognome + ", classe=" + classe + ", data=" + data + "]";
	}
	
}
```

### Example Excel File 

![Example Excel File](.//images//excel.png)

### Result 

```log
ExcelRow [nome=Mario, cognome=Arbola, classe=Quinta, data=19 Luglio 1996]
```

