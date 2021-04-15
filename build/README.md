# Compiling java source and rebuilding the OSGi bundle

Whenever either the java in `/src` or the jars in `/lib` change, the `lib-osgi.jar` file in the project root needs to be rebuilt.

`task.cfc` is a CommandBox task runner which will compile the java source into a jar, add it to the `/lib` directory and then re-build the `lib-osgi.jar` bundle in the root.

1. Edit `lib-osgi.mf` to increment the `Bundle-Version` and change the `Bundle-ClassPath` entries as required.
2. Copy the new Bundle-Version to the `osgiLibBundleVersion` property at the top of `Spreadsheet.cfc`.
3. Edit `task.cfc` if any of the `classpathDirectories` have changed
4. Open CommandBox in the this `build` directory and execute:
```
task run
```
## Note on `poi-ooxml` jars

The POI binary distribution includes two jar files:

`poi-ooxml-lite-VERSION.jar` and `poi-ooxml-full-VERSION.jar`

Only the *full* version is needed. Delete the lite version if copied to `/lib`