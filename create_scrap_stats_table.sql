INSERT INTO Apontamento ( Data, Hora, Operação, Máquina, Produzido, Refugado, Refugo )
SELECT Histórico.Data, Histórico.Hora, Histórico.Operação, Histórico.Máquina, Histórico.Produzido, Histórico.Refugado, FormatPercent([Refugado]/[Produzido],2) AS Refugo
FROM Histórico
WHERE (((Histórico.Data) Between #6/1/2022# And today()))
GROUP BY Histórico.Data, Histórico.Hora, Histórico.Operação, Histórico.Máquina, Histórico.Produzido, Histórico.Refugado, FormatPercent([Refugado]/[Produzido],2);
