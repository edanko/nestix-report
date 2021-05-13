namespace NestixReport
{
    public static class Db
    {
        public const string GetNxPathIds = @"SELECT
nxpath.nxname as name,
nxpath.nxpathid as id

FROM nxpath with (nolock) 
WHERE nxpath.nxname LIKE @name";

        public const string GetNxPathIdsForReport = @"select
nxpath.nxpathid
from nxpath with(nolock)
WHERE nxpath.nxname LIKE @name";

        public const string GetPartsFromNxPathId = @"select 
isnull(nxorderline.nxororderno,'') as CustOrderNo,
isnull(nxorderline.nxolsection,'') as Section,
isnull(nxproduct.nxprpartno,'') as PosMrpLineNo,
isnull(nxsheetpathdet.nxdetailcount*matpos.nxolordercount, 0) as DetailCount,
nxsheetpathdet.nxarea * nxorderline.nxolthick * PrPlate.nxprdensity as [Weight],
(nxsheetpathdet.nxarea * nxorderline.nxolthick * PrPlate.nxprdensity) * nxsheetpathdet.nxdetailcount as TotalWeight,
nxproduct.nxprthick,
nxproduct.nxprquality,
nxproduct.nxprlength as PartLength,
nxproduct.nxprwidth as PartWidth

from nxproduct with(nolock) 
inner join nxorderline with(nolock) on nxorderline.nxpartid = nxproduct.nxproductid 
inner join nxproduct as posmat with(nolock) on nxorderline.nxproductid = posmat.nxproductid
inner join nxsheetpathdet with(nolock) 
    inner join nxsheetpath with(nolock) on nxsheetpathdet.nxsheetpathid = nxsheetpath.nxsheetpathid
    inner join nxorderline as matpos with(nolock) on nxsheetpath.nxmatorderlineid = matpos.nxorderlineid
on nxsheetpathdet.nxorderlineid = nxorderline.nxorderlineid
  INNER JOIN nxorderline as matol with (nolock) on matol.nxorderlineid = nxsheetpath.nxmatorderlineid
  INNER JOIN nxproduct as PrPlate with (nolock) on PrPlate.nxproductid = matol.nxproductid

WHERE nxpathid = @pathid
order by PosMrpLineNo";


        public const string BatchNestingInfo = @"select 
nxpath.nxname,
dbo.NxDbGetNestBuildingSection(nxpath.nxpathid),
machine.name,
--nxsheetpath.nxused,
nxproduct.nxprthick,
nxproduct.nxprquality,
cast(round(ISNULL(nxorderline.nxollength, nxpath.nxmainlength),0) as nvarchar(20)) + 'x' + cast(round(ISNULL(nxorderline.nxolwidth, nxpath.nxmainheight),0) as nvarchar(20)),
ISNULL(nxpath.nxpathinfo, ''),
dbo.NxDbGetRemnantsForNesting(nxpath.nxpathid)

from nxpath with(nolock)

inner join machine with(nolock) on nxpath.nxmachineid = machine.machineid
inner join nxsheetpath with(nolock) on nxsheetpath.nxpathid = nxpath.nxpathid
inner join nxorderline with(nolock) on nxsheetpath.nxmatorderlineid = nxorderline.nxorderlineid
inner join nxproduct with(nolock) on nxorderline.nxproductid = nxproduct.nxproductid
left outer join nxvisual with(nolock) on nxpath.nxpathid = nxvisual.nxpathid

where nxpath.nxname LIKE @name
order by nxproduct.nxprthick asc, nxpath.nxname asc";


        public const string PickingList = @"select 
isnull(nxorderline.nxororderno,'') as OrderNo,
isnull(nxorderline.nxolsection,'') as Section,
isnull(nxproduct.nxprpartno,'') as PosNo,
isnull(nxsheetpathdet.nxdetailcount*matpos.nxolordercount, 0) as Count,
nxsheetpathdet.nxarea * nxorderline.nxolthick * PrPlate.nxprdensity as Weight,
(nxsheetpathdet.nxarea * nxorderline.nxolthick * PrPlate.nxprdensity) * nxsheetpathdet.nxdetailcount as TotalWeight,
nxproduct.nxprthick as Thickness,
nxproduct.nxprquality as Mat,
nxproduct.nxprlength as PartLength,
nxproduct.nxprwidth as PartWidth

from nxproduct with(nolock) 
inner join nxorderline with(nolock) on nxorderline.nxpartid = nxproduct.nxproductid 
inner join nxproduct as posmat with(nolock) on nxorderline.nxproductid = posmat.nxproductid
inner join nxsheetpathdet with(nolock) 
    inner join nxsheetpath with(nolock) on nxsheetpathdet.nxsheetpathid = nxsheetpath.nxsheetpathid
    inner join nxorderline as matpos with(nolock) on nxsheetpath.nxmatorderlineid = matpos.nxorderlineid
on nxsheetpathdet.nxorderlineid = nxorderline.nxorderlineid
  INNER JOIN nxorderline as matol with (nolock) on matol.nxorderlineid = nxsheetpath.nxmatorderlineid
  INNER JOIN nxproduct as PrPlate with (nolock) on PrPlate.nxproductid = matol.nxproductid

WHERE nxpathid = @pathid
--order by Section asc, PosNo asc";


        public const string PlatesInfo = @"select
nxpath.nxname,
machine.name,
nxsheetpath.nxused,
nxpath.nxmachtime,
nxproduct.nxprthick,
nxproduct.nxprquality,
round(ISNULL(nxorderline.nxollength, nxpath.nxmainlength), 0),
round(ISNULL(nxorderline.nxolwidth, nxpath.nxmainheight), 0),
nxproduct.nxprdensity * nxsheetpath.nxspnetarea * nxorderline.nxolthick

from nxpath with(nolock)

inner join machine with(nolock) on nxpath.nxmachineid = machine.machineid
inner join nxsheetpath with(nolock) on nxsheetpath.nxpathid = nxpath.nxpathid
inner join nxorderline with(nolock) on nxsheetpath.nxmatorderlineid = nxorderline.nxorderlineid
inner join nxproduct with(nolock) on nxorderline.nxproductid = nxproduct.nxproductid
left outer join nxvisual with(nolock) on nxpath.nxpathid = nxvisual.nxpathid
LEFT OUTER JOIN nxinventory with (nolock) ON nxorderline.nxorderlineid = nxinventory.nxinvmatorderlineid and nxinventory.nxorderlineid is null

where nxpath.nxname LIKE @name
order by nxproduct.nxprthick asc, nxpath.nxname asc";
    }
}