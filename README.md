# Cloud-ABAP2XLSX
This is a thinned/stripped down version of <a href="https://github.com/abap2xlsx/abap2xlsx">ABAP2XLSX</a> for ABAP Cloud systems. For complete functionality, onPrem users must still use the original ABAP2XLSX repo linked here.

Primary focus was given to the following changes:

1. Removed individual domains, data elements, structures, table types etc. Combined them all in one new Interface. This basically reduces the amount of ABAP objects created in the system.
2. Removed references to SALV*
3. Removed references to IXML replaced with Cloud equivalents
4. Removed references to MIME repo 
5. Removed references to Font 
6. Added a HTTP page with link to Demo programs
7. Removed CSV creation functionality
8. Removed App Server or Frontend references
9. Commented test classes, since they had many incompatible statements. Will uncomment them whenever each of them are fixed.
