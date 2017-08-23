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
