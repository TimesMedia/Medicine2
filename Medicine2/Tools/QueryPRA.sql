

 select 'Product' = z.[Name],  'Presentation' = a.[Name], 'NappiCode' = a.Nappi, 'Packaging' = c.[Name], c.NappiSuffix, 'DocumentGUID' = b.ID, z.MDRPage, 'Size' = Datalength(b.Content),
 'Format' = b.[Format] 
 from Product as z inner join SubProduct as a on z.ID = a.ProductId 
 inner join Document as b on a.DocumentID = b.ID
 inner join Pack as c on a.ID = c.SubProductId 
 where b.[Format] = '.docx' 
 and a.Nappi is not null and a.[Name] is not null and c.[Name] is not null and b.Id is not null and c.NappiSuffix is not null
 and z.Isrecycled = 0 and a.UseFreeStyle = 0 and DataLength(b.Content) > 0 
 --group by z.[Name], a.[Name], a.Nappi, c.NappiSuffix, c.Name, b.ID 
 order by a.[Name]

 select CONCAT('''', DocumentId,  '''',',')
from ProductClassification as a
where a.ClassificationID = 'A065E2AF-A806-415D-B210-ED9B3C257920'  -- MDR category




select a.ID, c.MDRPage, b.*, a.[Format],'Size' = Datalength(a.Content)
from Document as a inner join SubProduct as b on a.ID = b.DocumentId
inner join Product as c on c.Id = b.ProductID
where b.[Name] like 'Advantan%'

select CONCAT('''', DocumentId,  '''',',')
from ProductClassification as a
where a.ClassificationID = 'A065E2AF-A806-415D-B210-ED9B3C257920'