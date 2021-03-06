  SELECT V28.ARTICULO AS ARTICULO, V28.DESCRIP AS DESCRIPCION, V28.Ventas AS VENTAS_28_DIAS
  , ISNULL(V7.Ventas,0) AS VENTAS_7_DIAS 
  , V28.Existencia AS EXISTENCIA
  , CAST (ISNULL(EXISTENCIA*28/NULLIF(V28.Ventas,0),0) AS INT) AS DIAS_REST_EXIST FROM
 (SELECT pv.ARTICULO, sum(pv.CANTIDAD) as Ventas, prods.descrip, ISNULL(prods.existencia,0) as Existencia
  FROM [C:\MyBusinessDatabase\MyBusinessPOS2011.mdf].[dbo].[partvta] pv
  LEFT JOIN [C:\MyBusinessDatabase\MyBusinessPOS2011.mdf].[dbo].[prods] prods on pv.ARTICULO = prods.ARTICULO
  where pv.UsuFecha >= cast(current_timestamp-28 as datetime) and pv.UsuFecha <= cast(current_timestamp as datetime)
  group by pv.ARTICULO, prods.DESCRIP, prods.EXISTENCIA) V28
 LEFT JOIN   
 (SELECT sum(pv.CANTIDAD) as Ventas, 
 PV.ARTICULO
 FROM [C:\MyBusinessDatabase\MyBusinessPOS2011.mdf].[dbo].[partvta] pv
 where pv.UsuFecha >= cast(current_timestamp-7 as datetime) and pv.UsuFecha <= cast(current_timestamp as datetime)
 group by pv.ARTICULO) V7
 ON (V7.ARTICULO = V28.ARTICULO)
 where V28.Ventas <> 0
 order by V28.Ventas DESC, EXISTENCIA ASC;
 

  
  