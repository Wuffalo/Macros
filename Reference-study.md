Reference-study

VLOOKUP(Table245[[Index]:[Index]],Table1,MATCH(Table245[#Headers],Table1[#Headers],0),0)

INDEX(RefTable,MATCH(RowName,RowColumn,0),MATCH(ColName,ColumnRow,0))

wrong
INDEX(Table1,MATCH(Table245[[Index]:[Index]],(Table245[Index],0),MATCH(Table245[#Headers],Table1[#Headers],0))

right
INDEX(Table1,MATCH([@Index],[Index],0),MATCH(Table245[[#Headers],[Fox]],Table245[#Headers],0))

wrong
INDEX(Table1,MATCH(Table245[[Index]:[Index]],Table245[Index],0),MATCH(Table245[[#Headers],[Rabbit]],Table245[#Headers],0))

right
INDEX(RefTable,MATCH(WorkTable[[HIT]:[HIT]],RefTable[[HIT]:[HIT]],0),MATCH(WorkTable[#Headers],RefTable[#Headers],0))
