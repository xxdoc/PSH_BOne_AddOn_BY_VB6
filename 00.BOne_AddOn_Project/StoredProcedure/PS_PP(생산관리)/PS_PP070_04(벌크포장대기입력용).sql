USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_PP070_04]    Script Date: 04/14/2011 17:17:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : PP								    														*/
/*  Description    : 휘팅 벌크포장 대기				  															*/
/*  Create Date    : 2011.04.14                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
ALTER PROC [dbo].[PS_PP070_04]
--Create PROC [dbo].[PS_PP070_04]
(
	@ar_PP030No NVARCHAR(100)
	
)
AS

set @ar_PP030No = left(@ar_PP030No, charindex('-',@ar_PP030No ) - 1)


BEGIN
Create Table #Temp01 (
	PP030No		Nvarchar(20),
	OrdGbn		Nvarchar(20),
	OrdNum		Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	PP030HNo	Nvarchar(20),
	PP030MNo	Nvarchar(20),
	ORDRNo		Nvarchar(20),
	RDR1No		Nvarchar(20),
	ItemCode	Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	ItemName	Nvarchar(100),
	Selwt		Numeric(19,6),   
	JISURANG	Numeric(19,6),
	CpCode		Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	CPName		Nvarchar(30),
	PQty		Numeric(19,6),
	OutQty		Numeric(19,6),
	NQty		Numeric(19,6),
	GONGSENG	Numeric(19,6),
	DocDate		DateTime,
	JEGONG		Numeric(19,6),
	JEGONGWGT   Numeric(19,6),
	SEOULDE		Numeric(19,6),
	SEOULDEWt	Numeric(19,6),
	SEOULJE		Numeric(19,6),
	SEOULJEWt	Numeric(19,6),
	Sequence	Int	
)

--본사이동대기, 서울포장대기
Create Table #Temp02 (
	OrdNum		Nvarchar(20) Collate Korean_Wansung_Unicode_CI_AS,
	SEOULDE		Numeric(19,6),
	SEOULDEWt	Numeric(19,6),
	SEOULJE		Numeric(19,6),
	SEOULJEWt	Numeric(19,6)
)

Declare	@PP030No	Nvarchar(20),
		@OrdGbn		Nvarchar(20),
		@OrdNum		Nvarchar(20),
		@PP030HNo	Nvarchar(20),
		@PP030MNo	Nvarchar(20),
		@ORDRNo		Nvarchar(20),
		@RDR1No		Nvarchar(20),
		@ItemCode	Nvarchar(20),
		@ItemName	Nvarchar(100),
		@Selwt		Numeric(19,6),
		@JISURANG	Numeric(19,6),
		@CpCode		Nvarchar(20),
		@CPName		Nvarchar(30),
		@PQty		Numeric(19,6),
		@OutQty		Numeric(19,6),
		@NQty		Numeric(19,6),
		@GONGSENG	Numeric(19,6),
		@DocDate	DateTime,
		@JEGONG		Numeric(19,6),
		@SEOULDE	Numeric(19,6),
		@SEOULDEWt	Numeric(19,6),
		@SEOULJE	Numeric(19,6),
		@SEOULJEWt	Numeric(19,6),
		@Sequence	Int	



Declare	@BeforeOrdNum	Nvarchar(20), @BeforeGONGSENG Numeric(19,6)
Set @BeforeOrdNum = ''
Set @BeforeGONGSENG = 0
DECLARE CUR_1 CURSOR FOR
	Select	CONVERT(NVARCHAR,A.DocEntry) + '-' + CONVERT(NVARCHAR,B.LineId) AS PP030No
			,A.U_OrdGbn
			,A.U_OrdNum
			,A.DocEntry
			,B.LineId
			,A.U_SjNum
			,A.U_SjLine
			,CONVERT(NvarCHAR(60),A.U_ItemCode) AS ItemCode 
			,CONVERT(NVARCHAR(60),A.U_ItemName) AS ItemName
			,A.U_Selwt AS Selwt																		--지시수량
			,Convert(Numeric(19,2),A.U_SelWt) * Convert(Numeric(19,2),c.U_UnWeight) / 1000 As JISURANG
			,CONVERT(NVARCHAR(20),B.U_CpCode) AS CpCode
			,CONVERT(NVARCHAR(30),B.U_CpName) AS CpName
			,Isnull(Z.PQty, 0) As PQty
			,Isnull(Z.OutQty, 0) As OutQty
			,Isnull(Z.NQty, 0) As NQty
			,Isnull(Z.GONGSENG, 0) As GONGSENG
			,Isnull(Z.DocDate, 0) As DocDate
			,0 AS JEGONG
			,0 AS SEOULDE
			,0 AS SEOULDEWt
			,0 AS SEOULJE
			,0 AS SEOULJEWt
			,B.U_Sequence As Sequence
	  From	[@PS_PP030H] AS A 
			Inner Join [@PS_PP030M] AS B On A.DocEntry = B.DocEntry
			Inner Join [OITM] AS C On A.U_ItemCode = C.ItemCode
			Left  Join (Select	ISNULL(b.U_PP030HNo, '') As PP030HNO
								,ISNULL(b.U_Sequence, 0) As Sequence
								,ISNULL(b.U_CpCode, '') As CpCode
								,CASE When Max(a.U_OrdType) not in ('30', '60')  then Sum(IsNull(b.U_YQty, 0)) Else 0 End PQty
								,CASE When Max(a.U_OrdType) in ('30', '60') then Sum(IsNull(b.U_YQty, 0)) Else 0 End OutQty
								,Sum(IsNull(b.U_NQty, 0)) AS NQty															
								,SUM(ISNull(b.U_YQty, 0)) + Sum(IsNull(b.U_NQty, 0)) AS GONGSENG														
								,Max(a.U_DocDate) AS DocDate
						  From	[@PS_PP040H] a
								Inner Join [@PS_PP040L] b On a.DocEntry = b.DocEntry
								Inner Join [OITM] c On c.ItemCode = b.U_ItemCode
						 Where  a.Status = 'O'
						   
						Group by b.U_PP030HNo
								,b.U_Sequence 
								,b.U_CpCode
							 ) Z
							On  A.DocEntry = ISNULL(Z.PP030HNO, '') And ISNULL(B.U_Sequence, 0) = ISNULL(Z.Sequence, 0) and ISNULL(B.U_CpCode, '') = ISNULL(Z.CpCode, '')
	 WHERE	A.Status = 'O'
	   And  A.DocEntry = @ar_PP030No
	   
  
	Order by A.U_OrdNum, C.FrgnName, C.U_Spec1, C.U_Spec2, C.U_Spec3, C.U_Mark, C.U_ItemType,
			A.DocEntry, B.U_Sequence
OPEN CUR_1
FETCH NEXT FROM CUR_1 INTO 	@PP030No, @OrdGbn, @OrdNum, @PP030HNo, @PP030MNo, @ORDRNo, @RDR1No, @ItemCode, @ItemName, @Selwt, @JISURANG, @CpCode, @CPName, @PQty, @OutQty, 
							@NQty, @GONGSENG, @DocDate, @JEGONG, @SEOULDE, @SEOULDEWt, @SEOULJE, @SEOULJEWt, @Sequence
WHILE	@@FETCH_STATUS = 0
BEGIN				    
	
	If @BeforeOrdNum = @OrdNum Begin
		Set @JEGONG = @BeforeGONGSENG - @GONGSENG
	End	
	
	If @BeforeOrdNum <> @OrdNum Begin
		Set @BeforeGONGSENG = 0
		Set @JEGONG = 0
	End
	
	
	
	Insert #Temp01 values (	@PP030No, @OrdGbn, @OrdNum, @PP030HNo, @PP030MNo, @ORDRNo, @RDR1No, @ItemCode, @ItemName, @Selwt, @JISURANG, @CpCode, @CPName, @PQty, @OutQty, 
								@NQty, @GONGSENG, @DocDate, @JEGONG, 0, @SEOULDE, @SEOULDEWt, @SEOULJE, @SEOULJEWt, @Sequence)	
	Set @BeforeOrdNum = @OrdNum
	Set @BeforeGONGSENG = @GONGSENG - @NQty
	
	
	
FETCH NEXT FROM CUR_1 INTO 	@PP030No, @OrdGbn, @OrdNum, @PP030HNo, @PP030MNo, @ORDRNo, @RDR1No, @ItemCode, @ItemName, @Selwt, @JISURANG, @CpCode, @CPName, @PQty, @OutQty, 
							@NQty, @GONGSENG, @DocDate, @JEGONG, @SEOULDE, @SEOULDEWt, @SEOULJE, @SEOULJEWt, @Sequence
END	

CLOSE	CUR_1
DEALLOCATE CUR_1

Update #Temp01
   set JEGONGWGT = round((a.U_CpUnWt * #Temp01.JEGONG) / 1000,3)
   from [@PS_PP004H] a
  Where #Temp01.ItemCode = a.U_ItemCode and #Temp01.CpCode = a.U_CpCode

--공정단중 

--본사출하대기
Insert into #Temp02 ( OrdNum, SEOULDE, SEOULDEWt)
select b.U_OrdNum, Isnull(Sum(b.U_SelQty),0), Isnull(Sum(b.U_SelWt),0)
from [@ps_pp070H] a Inner Join [@ps_pp070L] b On a.DocEntry = b.DocEntry
where a.DocEntry = b.DocEntry
  and Isnull(b.U_MovDocNo,'') = ''
  and a.Canceled = 'N'
  And Exists (Select * From #Temp01 Where OrdNum = b.U_OrdNum)
 Group by b.U_OrdNum
  
 --서울포장대기
Insert into #Temp02 ( OrdNum, SEOULJE, SEOULJEWt )
--서울이동
select c.U_OrdNum,
	   Isnull(Sum(b.U_Qty),0),
	   Isnull(Sum(b.U_Weight),0)
	from [@PS_PP075H] a inner join [@PS_PP075L] b on a.docentry=b.docentry
						inner Join [@PS_PP070L] c On b.U_PP070No = convert(nvarchar(10),c.DocEntry) + '-' + convert(nvarchar(10),c.U_LineId)
	where a.Canceled = 'N'
	  And Exists (Select * From #Temp01 Where OrdNum = c.U_OrdNum)
   group by c.U_OrdNum	

Union All

select	b.U_OrdNum,
			Isnull(sum(a.U_NPkQty),0) * -1 As  Qty, 
			Isnull(sum(a.U_NPkWt),0) * -1 As  Wgt 
	from [@PS_PP077H] a Inner Join [@PS_PP070L] b On a.U_PP070No = b.DocEntry
	where a.Canceled = 'N'
	  And Exists (Select * From #Temp01 Where OrdNum = b.U_OrdNum)
group by b.U_OrdNum

Union all

select	c.U_OrdNum,
		Isnull(sum(b.U_PkQty - b.U_OPkQty),0) * -1 As Qty,
		Isnull(sum(b.U_PkWt - b.U_OPkWt),0) * -1 As Qty
	from [@PS_PP777H] a inner join [@PS_PP777L] b on a.docentry=b.docentry				
						Inner Join [@PS_PP070L] c On b.U_PP070No = convert(nvarchar(10),c.DocEntry) + '-' + convert(nvarchar(10),c.U_LineId)
	where a.Canceled = 'N'
	  And Exists (Select * From #Temp01 Where OrdNum = c.U_OrdNum)
group by c.U_OrdNum

--본사이동대기, 서울 포장대기 Update

 
 Update #Temp01
    set SEOULDE = Isnull(g.qty,0),
		SEOULDEWt = Isnull(g.wgt,0)
   from (Select a.Ordnum, qty = Isnull(Sum(a.SEOULDE),0), wgt = Isnull(Sum(a.SEOULDEWt),0)
		  from #Temp02 a
		 Group by a.Ordnum ) g
 Where #Temp01.OrdNum = g.OrdNum
   and #Temp01.CpCode = 'CP30112'
   
 Update #Temp01
    set SEOULJE = Isnull(g.qty,0),
		SEOULJEWt = Isnull(g.wgt,0)
   from (Select a.Ordnum, qty = Sum(a.SEOULJE), wgt = Sum(a.SEOULJEWt)
		  from #Temp02 a
		 Group by a.Ordnum
		 having Sum(a.SEOULJEWt) > 0 ) g
 Where #Temp01.OrdNum = g.OrdNum
   and #Temp01.CpCode = 'CP30112'


		
Select	a.PP030No,
		a.OrdGbn,
		a.OrdNum,
		OrdSub1 = '00',
		OrdSub2 = '000',
		a.PP030HNo As PP030HNo,
		a.PP030MNo As PP030MNo,
		a.ORDRNo As ORDRNo,
		a.RDR1No As RDR1No,
		BPLId = '1',
		a.ItemCode,
		a.ItemName,
		a.CpCode,
		a.Cpname,
		a.jegong - (Isnull(a.SEOULDE,0) + Isnull(a.SEOULJE,0))  As CpQty,
		a.jegongwgt - (Isnull(a.SEOULDEWt,0) + Isnull(a.SEOULJEWt,0)) As CpWt,
		0 AS SelQty,
		0 AS SelWt,			
		'' AS LineId
  From	#Temp01 a Inner join OITM b On a.ItemCode = b.ItemCode
 Where	OrdNum In (	SELECT	OrdNum
					  FROM	#Temp01
					 Where  JEGONG <> 0 Or SEOULDE <> 0 Or SEOULJE <> 0			  
				    Group by OrdNum	)
   and a.CpCode = 'CP30112'
Order by a.ItemName, a.OrdNum

End


--  EXEC [dbo].[PS_PP070_04] '1385-3'