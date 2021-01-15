Attribute VB_Name = "modMySQL"
'---------------------------------------------------------------------------------------------------------------
' ! Don't forget to add a reference to the Microsoft ActiveX Data Object Library in TOOLS -> REFERENCES !
' Author: Tomasz Kubiak t.kubiak@engineer.com
'---------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------
'Class factory - to make creating of the class easier:
'-----------------------------------------------------
Public Function rMySQL(Optional sConnectionString As String, Optional sUser As String, Optional sPassword As String, Optional lngOpts As Long) As clsMySQL
    Dim Result As clsMySQL
    Set Result = New clsMySQL
    Result.Connect sConnectionString, sUser, sPassword, lngOpts
    Set rMySQL = Result
End Function

'------------------------
'A FEW EXAMPLES OF USAGE:
'------------------------
Public Function TestMySQL()

    'PUT YOUR CONNECTION SETTINGS HERE TO LET EXAMPLES WORK:
        'You can use DSN:
        Const sMyConnectionString = "DSN=MyDSN"
        'OR define driver, server, port and user credentials directly.
        Const sMyConnectionString2 = "Driver={MySQL ODBC 5.3 Unicode Driver};Server=MyHostAddress;Port=3306;Database=INFORMATION_SCHEMA;UID=MyUserName;pwd=MyPassword"
    
    'it's not necessary to declare mysql object directly, but i like to do that because it makes intelli-sense available
        Dim MySQL As clsMySQL
    
    'EXAMPLE 1: Passing params as separate arguments.
        Debug.Print rMySQL(sMyConnectionString).Query("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_ROWS > ? LIMIT ?", 250, 10).GetString
    
    'EXAMPLE 2: Passing params as array. This time I'm using DNS, so I don't need to define User neither Passwors
        Debug.Print rMySQL(sMyConnectionString).Query("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_ROWS > ? LIMIT ?", Array(250, 10)).GetString
    
    'EXAMPLE 3: Define reusable clsMySQL object and use it for 3 times:
        Set MySQL = rMySQL(sMyConnectionString)
        MySQL.Query "CREATE OR REPLACE TEMPORARY TABLE tmpTest AS SELECT * FROM INFORMATION_SCHEMA.TABLES LIMIT 100"
        Debug.Print "Records affected: " & MySQL.RecordsAffected
        Debug.Print MySQL.Query("SELECT COUNT(*) FROM tmpTest").GetString
    
End Function
