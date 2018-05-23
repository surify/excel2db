# excel2db
Takes an excel file and loads into an sqlite3 database using python.

Description:

The below script takes
1) the path of a valid excel file and
2) the path of an existing or new sqlite3 database
from the user and transfers the table in the active sheet
with sheet's title as the database table name.

It currently assumes that the first row columns are headings

If there are any repeating column names, the user will be prompted
for a new name for that column in the middle of execution.

The source excel file will not be modified.

If you are giving the name of existing database, make sure that it
has no table with sheet's title in it(because it will be overwritten
if it exists).