SELECT Fluxo.Data, Fluxo.OPA, Fluxo.AN, Fluxo.Operação, Fluxo.Total AS Quantidade, Fluxo.Pirâmide, Fluxo.Parcial, [BD Agrupamentos].Agrupamento, [BD Fluxo por Tecnologia].FLUXO, [BD Fluxo por Tecnologia].TECNOLOGIA, Date()-[Data] AS WIP INTO [Report WIP]
FROM Fluxo, [BD Agrupamentos], [BD Fluxo por Tecnologia]
WHERE (((Fluxo.AN)=[BD Fluxo por Tecnologia]!AN) And (([BD Agrupamentos].AN)=Fluxo!AN) And (([BD Fluxo por Tecnologia].AN)=Fluxo!AN));
