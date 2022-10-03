SELECT
    ArcDateTime, ProgramName, RepeatID,
    SheetName, PartName, WONumber,
    QtyInProcess, NestedArea
FROM
    PIPArchive
WHERE 
    PartName NOT LIKE '11%'
AND
    PartName NOT LIKE '12%'
AND
    TransType='SN102'
ORDER BY
    ArcDateTime DESC