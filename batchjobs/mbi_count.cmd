:: RUN SOME SQL QUERY CODE TO EXPORT THE NUMBER OF BLOCKS INTO A TEXT FILE WE MAIL NEXT
sqlcmd -S "." -d XCA_DB -Q "Declare @Count Int; Declare @Text varchar(500); Declare @Stamp varchar(50); Set @Stamp=(GetDate()); Set @Count=(Select Count (*) From xca_MBI Where Status In ('A','I')); Set @Text='There were ' + Cast(@Count as varchar) + ' MBI Blocks assigned as of ' + Cast(@Stamp as varchar) + '.  Do not forget to update the Jan 01 value.'; Print @Text" > D:\CNA\Internal\MBICount.txt

:: MAIL THE TEXT FILE WITH BLAT
blat "D:\CNA\internal\MBICount.txt" -subject "MBI Count"  -to "browng@leidos.ca"  -cc "walshkel@leidos.ca"  -mailfrom "database@leidos.ca" -from "database@leidos.ca" -replyto "browng@leidos.ca" -u "admin" -pw "&Keep@Bay" -server "192.168.10.51"

