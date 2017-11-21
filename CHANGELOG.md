# 2.0.4 Manually shiftRowsAndMergedRegions

POI breaks merged regions using shiftRows

# 2.0.3 Helper to get all static MergeFields of a Class

The typing suppport gets partially lost when using an Enum that implents MergeFieldProvider.
The recommended method to declare MergeFields is to use a final Class with static MergeFields for each Excel template.

The new helper method `MergeField.getStaticMergeFields(YourStaticClass.class)` can be used to get all these static
MergeFields inorder to pass them to the ExcelMerger.
# 2.0.2 no changes

# 2.0.1 Naming conventions

We dropped 'lib' in the `groupId` and `artifactId` name to comply with new naming conventions.

Use the following dependency from now on:

```xml
<dependency>
	<groupId>ch.dvbern.oss.excelmerger</groupId>
	<artifactId>excelmerger-impl</artifactId>
	<version>(NEWEST_VERSION)</version>
</dependency>
```

# 2.0.0 Typing support

## License change
We are open sourcing! This library is now available under the Apache License, Version 2.0 

## POI Update
Upgrade from 3.15 to 3.16

## Breaking change
Deprecated code has been removed.

MergeFields and Converters are now using property typings, such that you cannot set mistakingly
the wrong merge value.

If you used excel Merger before, you will have to adjust your MergeField enums. 
You can Use the following regular expressions to adjust to the new api.

```
(\w+)\((\w+), Type.SIMPLE\)
$1(new SimpleMergeField<>("$1", $2))
```

```
(\w+)\((\w+), Type.REPEAT_COL\)
$1(new RepeatColMergeField<>("$1", $2))
```

```
(\w+)\((\w+), Type.REPEAT_VAL\)
$1(new RepeatValMergeField<>("$1", $2))
```

```
(\w+)\((\w+), Type.REPEAT_ROW\)
$1(new RepeatRowMergeField("$1"))
```

```
enum (\w+) implements MergeField
enum $1 implements MergeFieldProvider
```

```
import static ch\.dvbern\.lib\.excelmerger\.StandardConverters\.(\w+);
import static ch\.dvbern\.oss\.lib\.excelmerger\.converters\.StandardConverters\.$1;
```

```
import ch\.dvbern\.lib\.excelmerger\.MergeField;
import ch\.dvbern\.oss\.lib\.excelmerger\.mergefields\.MergeField;
```

```
ch\.dvbern\.lib\.excelmerger
ch\.dvbern\.oss\.lib\.excelmerger
```

Please adjust your enum constructors and getters manually.
