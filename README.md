# spreadsheet-copy
Extension to the Apache Java POI library to allow copying and merging of spreadsheets

## Getting Started
**Prerequisite**\
Before you get started, the Apache POI library must be downloaded and referenced in\
CLASSPATH.\
You can download the latest version here: https://poi.apache.org/download.html

**Download & Build**
- Clone this Git repository to your local machine
- Make the script directory your current working directory (cd script)
- Run the jar-build.sh command to build the jar file
- Reference the spreadsheet-copy-fourjs.jar file in your CLASSPATH

### Generate JWT Token 

The following code can be used to generate JWT token in Java

**Java Implementation**
```java

import com.fourjs.jwt.JWebToken;

String subject = "John";
String[] audience = new String[]{"Admin", "Accounting", "Payroll"};
int expireMinutes = 30;
String secretKey = "ThisIsMySpecialSecretKey:OU812!";

JWebToken token = JWebToken.CreateToken(subject, audience, expireMinutes, secretKey);

String bearerToken = token.toString();

```

**Genero Implementation**
```genero
IMPORT JAVA com.fourjs.jwt.JWebToken
IMPORT JAVA java.lang.String
IMPORT JAVA java.lang.Object

PUBLIC TYPE JavaString java.lang.String
PUBLIC TYPE JavaStringArray ARRAY[] OF JavaString
PUBLIC TYPE JavaObjectArray ARRAY[] OF java.lang.Object

PUBLIC TYPE TWebToken com.fourjs.jwt.JWebToken

FUNCTION buildToken() RETURNS ()
   DEFINE token TWebToken
   DEFINE permissions DYNAMIC ARRAY OF STRING = ["Admin", "Accounting", "Payroll"];
   DEFINE javaPermissions JavaStringArray
   DEFINE idx INTEGER
   DEFINE expireMinutes INTEGER
   DEFINE secretKey STRING
   DEFINE subject STRING
   DEFINE bearerToken STRING

   LET subject = "John"
   LET expireMinutes = 30
   LET secretKey = "ThisIsMySpecialSecretKey:OU812!"
   LET javaPermissions = JavaStringArray.create(permissions.getLength())
   FOR idx = 1 TO permissions.getLength()
      LET javaPermissions[idx] = permissions[idx]
   END FOR

### Credits
This code was modified from the original implementation done here:\
http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
