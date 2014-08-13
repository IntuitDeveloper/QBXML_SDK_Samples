@echo off
SETLOCAL
if "%TOMCAT_HOME%"=="" set TOMCAT_HOME=c:\work\install\Tomcat-4.1.31
if "%AXIS_HOME%"=="" set AXIS_HOME=\root\axis-1_2_1
if "%JAVA_HOME%"=="" set JAVA_HOME=\usr\java\j2sdk1.4.2_09
set CLASSPATH=%TOMCAT_HOME%\bin\bootstrap.jar;%JAVA_HOME%\lib\tools.jar;%JAVA_HOME%\lib

set AXIS_LIB=%AXIS_HOME%\WEB-INF\lib
set AXISCLASSPATH=%AXIS_LIB%\axis.jar;%AXIS_LIB%\axis-ant.jar;%AXIS_LIB%\commons-discovery-0.2.jar;%AXIS_LIB%\commons-logging-1.0.4.jar;%AXIS_LIB%\jaxrpc.jar;%AXIS_LIB%\log4j-1.2.8.jar;%AXIS_LIB%\saaj.jar;%AXIS_LIB%\wsdl4j-1.5.1.jar;%AXIS_LIB%\xml-apis.jar;%AXIS_LIB%\xercesImpl.jar

%JAVA_HOME%\bin\javac -classpath %CLASSPATH%;%AXISCLASSPATH%;. %1
ENDLOCAL
