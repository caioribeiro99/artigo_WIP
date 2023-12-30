SELECT Encerrado.Data, Sum(Encerrado.Total) AS SomaDeTotal, Encerrado.Operação
FROM Encerrado
WHERE (((Encerrado.Data)=IIf(Weekday(Date())=2,Date()-3,Date()-1)))
GROUP BY Encerrado.Data, Encerrado.Operação;
