# Quarantainenet-export
Export all endpoints in quarantainenet NAC solution by walking the webinterface

There is probably a way better way to do this. But this is what we did to export the endpoint database when migrating to clearpass (10.000+ endpoints). The script does get confused on the last page so check the manualy.

```Navigates to the webinterface
Logs in
*Use the old non ajex interface*
Walks all the pages and greps: 
MAC, AccessGroup, RealName, Description, Registered by, Time of registration and expire date.
