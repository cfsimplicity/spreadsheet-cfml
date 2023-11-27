# Compiling java source and rebuilding the OSGi bundle

Whenever either the java in `/src` or the jars in `/lib` change, the `lib-osgi.jar` file in the project root needs to be rebuilt.

`task.cfc` is a CommandBox task runner which will compile the java source into a jar, add it to the `/lib` directory and then re-build the `lib-osgi.jar` bundle in the root.

1. Edit `lib-osgi.mf` to increment the `Bundle-Version` and change the `Bundle-ClassPath` entries as required.
2. Copy the new Bundle-Version to the `osgiLibBundleVersion` property at the top of `Spreadsheet.cfc`.
3. Edit `task.cfc` if any of the `classpathDirectories` jar versions have changed
4. Open CommandBox in the this `build` directory and execute:
```
task run
```
## POI binaries

From version 5.2.4 a binary distribution (zip) is no longer provided. This means the POI, POI OOXML and POI OOXML-FULL jars need to be downloaded individually from

https://repo1.maven.org/maven2/org/apache/poi/

Dependencies and versions are given in the .pom files in the /poi, /poi-ooxml and /poi-ooxml-full directories. These jars can be downloaded by searching:

https://mvnrepository.com

Here is the current list of required jars:

* commons-codec
	https://mvnrepository.com/artifact/commons-codec/commons-codec
* commons-collections4
	https://mvnrepository.com/artifact/org.apache.commons/commons-collections4
* commons-compress
	https://mvnrepository.com/artifact/org.apache.commons/commons-compress
* commons-io
	https://mvnrepository.com/artifact/commons-io/commons-io
* commons-math3
	https://mvnrepository.com/artifact/org.apache.commons/commons-math3
* log4j-api
	https://mvnrepository.com/artifact/org.apache.logging.log4j/log4j-api
* poi
	https://repo1.maven.org/maven2/org/apache/poi/poi/
* poi-ooxml
	https://repo1.maven.org/maven2/org/apache/poi/poi-ooxml/
* poi-ooxml-full
	https://repo1.maven.org/maven2/org/apache/poi/poi-ooxml-full
* SparseBitSet
	https://mvnrepository.com/artifact/com.zaxxer/SparseBitSet
* xmlbeans
	https://mvnrepository.com/artifact/org.apache.xmlbeans/xmlbeans

## Other dependency binaries

The following jars are not POI dependencies but are required by the Spreadsheet Library:

* commons-csv
	https://mvnrepository.com/artifact/org.apache.commons/commons-csv
* excel-streaming-reader
	https://mvnrepository.com/artifact/com.github.pjfanning/excel-streaming-reader
* slf4j-api
	https://mvnrepository.com/artifact/org.slf4j/slf4j-api

All available from https://mvnrepository.com