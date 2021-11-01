# sql_extract_stored_proc

![Banner](./banner.svg)

> VBS script for extracting every stored procedures from SQL Server databases into flat files on disk.

## Table of Contents

- [Description](#description)
- [Install](#install)
- [Usage](#usage)
- [Author](#author)
- [License](#license)

## Description

Connect to a SQL Server database, obtain the list of
stored procedures (USPs) in that db (process all schemas), get
the code in these stored procs and save them as text files (.md)

At the end, we'll have as many files as there are stored procs
in the database. One text file by stored proc.

Files will be saved under the /results folder.

Running this script against a SQL Server DB will take a local
copy of your USPs : you can then take a backup of them easily.

NOTE : The user should have enough permissions on SQL Server side
for retrieving the code of the stored procedure. This is never the
case of a "simple" user and requires advanced permissions. So; if
generated files are empty, first check user's permissions (or directly
use an "admin" user to check if it's better).

NOTE : files are stored as a .md file to make easier to include them into [marknotes](https://github.com/cavo789/marknotes), a documentation tool.

## Install

Just get a copy of the sql_extract_stored_proc.vbs file and store it onto your computer.

## Usage

You'll need to provide the script with four variables :

- The SQL Server name (f.i. `myServer`)
- The name of the database (f.i. `dbOfMine`)
- The user to use for the connection (f.i. `userAdmin`)
- The password for this user (f.i. `my$ecret`)

You can pass these parameters as command line arguments

```
cscript sql_extract_stored_proc.vbs "myServer" "dbOfMine" "userAdmin" "my$ecret"
```

(so you can reuse this script for other databases)

or by updating, in the source code, the constants that you can find at the top of the file. If you do that, you don't need to specify credentials anymore and you can just fire the script like :

```
cscript sql_extract_stored_proc.vbs
```

When the script is done, you'll get a folder called `results´ with one .md file by stored procedure. The .md file is a markdown file format with the source code of the stored procedure.

## Author

Christophe Avonture

## Contribute

PRs not accepted.

## License

[MIT](LICENSE)
