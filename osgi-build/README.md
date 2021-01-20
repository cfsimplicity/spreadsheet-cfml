# Rebuilding the OSGi bundle

Whenever one or more of the jar files in the `/lib` directory change, the `lib-osgi.jar` file in the project root needs to be rebuilt.

1. Edit `lib-osgi.mf` to increment the `Bundle-Version` and change the `Bundle-ClassPath` and `Export-Package` entries as required.
2. Copy the new Bundle-Version to the `osgiLibBundleVersion` property at the top of `Spreadsheet.cfc`.
3. Open a command in the project root directory and execute:
```
jar -cvfm lib-osgi.jar osgi-build/lib-osgi.mf -C lib/ .
```
## Note on `poi-ooxml` jars

The POI binary distribution includes two jar files:

`poi-ooxml-lite-VERSION.jar` and `poi-ooxml-full-VERSION.jar`

Only the *lite* version is needed. Delete the full version if copied to `/lib`