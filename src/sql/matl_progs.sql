SELECT
    PIPArchive.ProgramName,
    PIPArchive.WONumber,
    PIPArchive.PartName,
    PIPArchive.QtyInProcess,
    PIPArchive.NestedArea,
    StockHistory.SheetName,
    StockHistory.PrimeCode,
    StockHistory.Mill
FROM PIPArchive
INNER JOIN StockHistory
    ON StockHistory.ProgramName=PIPArchive.ProgramName
WHERE
    StockHistory.PrimeCode=?
AND
    PIPArchive.TransType='SN102'
