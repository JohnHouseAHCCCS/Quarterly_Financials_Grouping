How to update:
    1. Download a set of quarterly financials from here:
        https://ahcccs.sharepoint.com/sites/DHCMSharedFinandRI/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FDHCMSharedFinandRI%2FShared%20Documents%2FFIN%2FHealth%20Plan%20Summaries%2FQuarterly%20Filings&viewid=958924a2%2D0d2d%2D4b36%2Db47c%2D5dc609c0f471
    2. Extract all, then drag the resulting directory that contains all the quarter folders over "move_files.py"
    3. Run data_extraction.py (this takes a while).
    4. Check data_extraction.log for duplicates, then delete appropriate files
    5. Repeat steps 3-4 until there are no duplicates
    6. Run data_aggregation.py
    7. Results are in /Output/Results
    8. Move these files to "T:\Quarterly Financials\Statements of Activities" after archiving the previous set