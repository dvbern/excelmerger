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

``
(\w+)\((\w+), Type.SIMPLE\)
$1(new SimpleMergeField<>("$1", $2))
``

``
(\w+)\((\w+), Type.REPEAT_COL\)
$1(new RepeatColMergeField<>("$1", $2))
``

``
(\w+)\((\w+), Type.REPEAT_VAL\)
$1(new RepeatValMergeField<>("$1", $2))
``

``
(\w+)\((\w+), Type.REPEAT_ROW\)
$1(new RepeatRowMergeField("$1"))
``

``
enum (\w+) implements MergeField
enum $1 implements MergeFieldProvider
``

``
import static ch\.dvbern\.lib\.excelmerger\.StandardConverters\.(\w+);
import static ch\.dvbern\.oss\.lib\.excelmerger\.converters\.StandardConverters\.$1;
``

``
import ch\.dvbern\.lib\.excelmerger\.MergeField;
import ch\.dvbern\.oss\.lib\.excelmerger\.mergefields\.MergeField;
``

``
ch\.dvbern\.lib\.excelmerger
ch\.dvbern\.oss\.lib\.excelmerger
``

Please adjust your enum constructors and getters manually.

# 1.1.1 PERCENT_CONVERTER scale issue
- It was possible, that the resulting percentage value was wrongly converted, for
instance when the scale of the input BigDecimal was 0.

# 1.1.0 FZL (kurstool) integration (2017-03-17)
  
## Changes
- INTEGER_CONVERTER has been deprecated because it is actually a LONG_CONVERTER. 
Please migrate to LONG_CONVERTER.

## New Features
- Template-Patters may contain underscore "_"
- **MergeField** 
  - PAGE_BREAK - can be added in a template to insert a page break
- **Converters**
  - STRING_COLORED_CONVERTER - write a string value in the supplied colour
  - DATE_WITH_DAY_CONVERTER - prefixes the default date format with the weekday ("EE, dd.MM.yyyy")
  - LONG_CONVERTER - superseeds INTEGER_CONVERTER
  
## Bugfixes
- slf4j-log4j12 should only be included in test scope

# 1.0.0 KitAdmin Excel Merger (2017-03-17)
