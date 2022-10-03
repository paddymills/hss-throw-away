SELECT
    PIPArchive.PartName AS part,
    PIPArchive.WONumber AS wo,
    PIPArchive.QtyInProcess AS qty,
    Part.Data1 AS job,
    Part.Data2 AS shipment
FROM PIPArchive
INNER JOIN (
    SELECT PartName, WONumber, Data1, Data2 FROM Part
    UNION
    SELECT PartName, WONumber, Data1, Data2 FROM PartArchive
) AS Part
    ON PIPArchive.PartName=Part.PartName
    AND PIPArchive.WONumber=Part.WONumber
WHERE
    PIPArchive.WONumber IN (
        SELECT WONumber FROM Wo WHERE WONumber LIKE '%-%-%'
        UNION
        SELECT WONumber FROM WOArchive WHERE WONumber LIKE '%-%-%'
        UNION
        SELECT 'FARO' AS WONumber
    )
AND
    PIPArchive.TransType='SN102'
AND
    PIPArchive.ArcDateTime > ?
ORDER BY PIPArchive.PartName DESC