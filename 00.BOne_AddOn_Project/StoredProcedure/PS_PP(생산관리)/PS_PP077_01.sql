USE [PSHDB_START]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP077_01]    Script Date: 11/24/2010 20:48:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅서울포장등록																				*/
/*  ALTER  Date    : 2010.10.26																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP077_01]
ALTER     PROC [dbo].[PS_PP077_01]
(
  @MovDocNo		as nvarchar(15),
  @FrDate		as datetime,
  @ToDate		as datetime,
  @ItemCode		as nvarchar(30),
  @Status		as nvarchar(1)
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
--select * from [@PS_PP070H]
--select * from [@PS_PP070L]

/*
update [@PS_PP070L]
set u_itemcode='11010003'
where docentry=1
and lineid=2

update [@PS_PP075L]
set u_itemcode='11010003', u_size='41.28'
where u_pp070no='1-2'
*/
-----------------------------------------------------------------------------------------------------------------------------------------
/*
select	b.U_MovDocNo '이동문서번호',
		a.DocNum	 '벌크포장번호',
		b.U_ItemCode '제품코드',
		b.U_ItemName '제품이름',
		c.U_Size	 '규격',
		d.Name       '인증기호',
		count(b.U_ItemCode) '항목수',
		sum(b.U_SelQty)		'총수량',
		sum(b.U_SelWt)	    '총중량'
from [@PS_PP070H] a inner join [@PS_PP070L] b on a.docentry=b.docentry
				    left  join [OITM] c on b.U_ItemCode=c.ItemCode
					left  join [@PSH_MARK] d on c.U_Mark=d.Code
					
--where b.U_MovDocNo like @MovDocNo				
--and	  b.U_ItemCode like @ItemCode
where	isnull(b.U_MovDocNo,'') <> ''
group by b.U_MovDocNo,a.DocNum,b.U_ItemCode,b.U_ItemName,c.U_Size,d.Name
*/					
										
--select * from [@PS_PP075H]
--select * from [@PS_PP075L]

select	a.U_MovDocNo,
		left(b.U_PP070No,charindex('-',b.U_PP070No)-1) PP070No,
		b.U_ItemCode,
		b.U_ItemName,
		b.U_Size,
		b.U_Mark,
		count(b.U_ItemCode) Cnt,
		sum(b.U_Qty) Qty, 
		sum(b.U_Weight) Weight
into #PS_PP077_01
from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry				
where a.Canceled <> 'Y'
and	  a.U_MovDocNo like @MovDocNo				
and	  b.U_ItemCode like @ItemCode
and   a.U_RegiDate between @FrDate and @ToDate
group by a.U_MovDocNo, 
		 left(b.U_PP070No,charindex('-',b.U_PP070No)-1),
		 b.U_ItemCode,b.U_ItemName,
		 b.U_Size,
		 b.U_Mark					
-----------------------------------------------------------------------------------------------------------------------------------------
--select * from #PS_PP077_01
-----------------------------------------------------------------------------------------------------------------------------------------
if @Status='1' begin --포장대기

	select	a.U_MovDocNo'이동문서번호',
			a.PP070No'벌크포장번호',
			a.U_ItemCode'제품코드',
			a.U_ItemName'제품이름',
			a.U_Size'규격',
			a.U_Mark'인증기호',
			a.Cnt'항목수',
			a.Qty'총수량', 
			a.Weight'총중량',
			isnull(b.NPkQty,0)'기포장수량',
			isnull(b.NPkWt,0)'기포장중량'
	from #PS_PP077_01 a left  join (
									select U_MovDocNo, U_ItemCode, U_PP070No, sum(isnull(U_NPkQty,0)) NPkQty, sum(isnull(U_NPkWt,0)) NPkWt
									from [@PS_PP077H]
									group by U_MovDocNo, U_ItemCode, U_PP070No	 
									) b on a.U_MovDocNo=b.U_MovDocNo and a.U_ItemCode=b.U_ItemCode and a.PP070No=b.U_PP070No
	where a.Weight > isnull(b.NPkWt,0)
	or    isnull(b.NPkWt,0) = 0
		 
end else begin --포장완료

	select	a.U_MovDocNo'이동문서번호',
			a.PP070No'벌크포장번호',
			a.U_ItemCode'제품코드',
			a.U_ItemName'제품이름',
			a.U_Size'규격',
			a.U_Mark'인증기호',
			a.Cnt'항목수',
			a.Qty'총수량', 
			a.Weight'총중량',
			isnull(b.NPkQty,0)'기포장수량',
			isnull(b.NPkWt,0)'기포장중량'
	from #PS_PP077_01 a left  join (
									select U_MovDocNo, U_ItemCode, U_PP070No, sum(isnull(U_NPkQty,0)) NPkQty, sum(isnull(U_NPkWt,0)) NPkWt
									from [@PS_PP077H]
									group by U_MovDocNo, U_ItemCode, U_PP070No	 
									) b on a.U_MovDocNo=b.U_MovDocNo and a.U_ItemCode=b.U_ItemCode and a.PP070No=b.U_PP070No
	where a.Weight <= isnull(b.NPkWt,0)
	and    isnull(b.NPkWt,0) <> 0

end
-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_PP077_01 '%','20101101','20101124','%','2'
--EXEC PS_PP077_01 '%','20101101','20101124','%','1'