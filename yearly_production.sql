SELECT StrConv(MonthName(Month([Data])),1) AS Mês, Sum(Encerrado.Total) AS SomaDeTotal
FROM Encerrado
WHERE ((Year([Encerrado]![Data])=Year(Date())))
GROUP BY StrConv(MonthName(Month([Data])),1);
