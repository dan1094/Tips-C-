
//Lista de objetos contenidos 
//porq es sacar todo lo que tiene el conjunto de A pero que no lo tiene B
//teoria de conjuntos (A - B)
//LEFT JOIN
var diferentes = (from r in lista1 
                                   join p in Lista2 on r.Identificador equals p.Identificador
                                   into temporal
                                   from d in temporal.DefaultIfEmpty()
                                   where d == null
                                   select r).ToList();