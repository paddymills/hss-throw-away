SELECT
    PIP.ArcDateTime, PIP.ProgramName, PIP.SheetName,
    PIP.PartName, PIP.WONumber, PIP.QtyInProcess,
    PIP.NestedArea
FROM PIPArchive AS PIP
    INNER JOIN Part
        ON
            Part.PartName=PIP.PartName
        AND
            Part.WONumber=PIP.WONumber
WHERE
    Part.Data5 = ''
AND
    PIP.TransType='SN102'
AND
    PIP.ArcDateTime > '2020-09-01'
ORDER BY PIP.ArcDateTime DESC