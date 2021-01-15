# MySQL4ADODB
Simple VBA class to make work with VBA and MySQL (especially passing query parameters) much easier and more handy.

### EXAMPLE #1: Passing params as separate arguments
```
Debug.Print rMySQL("DSN=MyDSN").Query("SELECT * FROM MyTable WHERE Id > ? LIMIT ?", 250, 10).GetString
```
### EXAMPLE #2: Passing params as array:
```
Debug.Print rMySQL("DSN=MyDSN").Query("SELECT * FROM MyTable WHERE Id > ? LIMIT ?", Array(250, 10)).GetString
```
### EXAMPLE #3: Define reusable clsMySQL object:
    Set MySQL = rMySQL("DSN=MyDSN")
    MySQL.Query "CREATE OR REPLACE TEMPORARY TABLE tmpTest AS SELECT * FROM Mytable WHERE Id<200 LIMIT 15"
    Debug.Print "Records affected: " & MySQL.RecordsAffected
    Debug.Print MySQL.Query("SELECT COUNT(*) FROM tmpTest").GetString
