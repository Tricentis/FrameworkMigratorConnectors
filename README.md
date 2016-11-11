# FrameworkMigratorConnectors

This is a sample project to explain how to write Framework Migration Connector for Tosca Testsuite using TCMigrationAPI.

Framework Migration Connectors contain Frameowrk specific migration mapping logic that are used to convert Framework artifacts to Tosca Business Objects.
It uses the generic functionalities provided by TCMigrationAPI and packaged as a Tosca AddOn task (context menu task).

**The Application under Test for the Sample Framework is the Vehicle Insurance Application provided by [Tricentis](http://www.tricentis.com/) and can be accessed [here](http://www.toscatrial.com/sampleapp/)

**The FrameworkMigratorConnector sample projects are built on Tosca Testsuite 10.0

## How you can use it

1. Download the Sample FrameworkMigratorConnector solution.
2. Download the TCMigrationAPI assembly binary [here](https://github.com/Tricentis/FrameworkMigratorAPI/tree/master/TCMigrationAPI/bin)
3. Reference the downloaded TCMigrationAPI assembly into FrameworkMigratorConnector project.
4. The other dependencies should be auto-referenced, if Tosca Testsuite is installed in the system.
5. Upon successful build, the binaries should be copied to %TRICENTIS_HOME% automatically.
6. Open Tosca, access the context menu task from workspace root and upload the sample artifacts provided with the solution.
