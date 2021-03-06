USE [PSHDB]
GO
/****** Object:  StoredProcedure [dbo].[PS_SD_SMARTBILL]    Script Date: 03/23/2011 13:02:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/******************************************************************************************************************/
/*  Module         : SD								    														*/
/*  Description    : 스마트빌 세금계산서																		*/
/*  Create Date    : 2011.03.23                                                                                   */
/*  Modified Date  :										       													*/
/*  Creator        : N.G.Y																						*/
/*  Company        : Poongsan Holdings																			*/
/******************************************************************************************************************/
ALTER PROC [dbo].[PS_SD_SMARTBILL]
--Create PROC [dbo].[PS_SD_SMARTBILL]
(
	@BPLId      NVARCHAR(1),
	@DocDateFr  datetime,
	@DocDateTo  datetime
   
)
AS

BEGIN
Create Table #Temp01 (
	cnt			Nvarchar(6),
	VatRegNum	Nvarchar(20), --사업자번호
	Email		Nvarchar(100), --Email
	taxgubun	Char(1),	  --세금계산서 구분
	Seqn		Nvarchar(20), --세금계산서 번호(송장번호)
	ymd			Nvarchar(10), --계산서 일자
	amt			Numeric(19,0),--계산서 공급가
	vat			Numeric(19,0),--계산서 부가세
	tot			Numeric(19,0),--계산서총계
	mmdd		Char(4),	  --계산서년월
	ItemCode	Nvarchar(20), --제품코드
	ItemName	Nvarchar(200),--제품명
	size	    Nvarchar(200),--규격
	Quantity	Numeric(19,0),--수량
	Price		Numeric(19,6),--단가
	damt		Numeric(19,0),--공급가
	dvat		Numeric(19,0),--부가세
	bigo		Numeric(19,3),--중량
	CardName    Nvarchar(100),--거래처
	DocEntry	Nvarchar(20), --송장번호
	LineNum		Nvarchar(20)  --송장순번
 )

Declare	@DocEntry	Nvarchar(20), @LineNum	Nvarchar(20)
Declare @cnt		int
Set @cnt = 1

Declare	@BeforeDoc	Nvarchar(20)
Set @BeforeDoc = ''

Insert Into #Temp01
select Cnt = '',
	   i.VatRegNum,
       i.E_Mail,
       taxgubun = (CASE WHEN a.VatGroup = 'A3' then 'P' /* 로칼=영세 */
                   ELSE 'T'   /* 내수=과세 */
                   END),
       seqn     = 'AR' + convert(nvarchar(10),a.DocEntry),
       ymd      = (select Convert(char(8),DocDate,112) From OINV where DocEntry = a.DocEntry ),
       amt      = a.GrosProfSy,
	   vat      = a.VatSumSy,
       tot      = a.DocTotalSy,
       mmdd     = (select right(Convert(Char(8),DocDate,112) ,4) From OINV where DocEntry = a.DocEntry ),
       ItemCode   = a.ItemCode,
       ItemName   = c.FrgnName + (CASE WHEN c.ItmsGrpCod = '101' THEN ' [상품]'
                              ELSE ''
                              END),
       spec     = isnull(Convert(Nvarchar(100),c.U_Size),'') + ' ' +isnull(Convert(Nvarchar(10),f.Name),''),
       qty      = a.Quantity,
       price    = a.price,
	   damt     = a.LineTotal,
	   dvat     = a.VatSum,
       dbigo    = round((a.Quantity * c.U_UnWeight) / 1000,3),
       i.CardName,
       DocEntry = 'AR' + convert(nvarchar(10),a.DocEntry),
       LineNum = Convert(nvarchar(10),a.LineNum)
From (

Select t1.DocEntry,
	   t1.LineNum,
	   t1.CardCode,
	   t1.VatGroup,
	   GrosProfSy = sum(t1.GrosProfSy),
	   VatSumSy = sum(t1.VatSumSy),
	   DocTotalSy = sum(t1.DocTotalSy),
	   ItemCode = t1.ItemCode,
	   Quantity = sum(t1.Quantity),
	   price = sum(t1.Price),
	   LineTotal = Sum(t1.LineTotal),
	   VatSum = sum(t1.VatSum)
From (
Select DocEntry      = a.DocEntry,
	   LineNum  = b.LineNum,
	   CardCode  = a.CardCode,
       GrosProfSy = a.GrosProfSy,
	   VatSumSy      = a.VatSumSy,
       DocTotalSy = a.DocTotalSy,
       ItemCode   = b.ItemCode,
       Quantity = b.Quantity,
       price    = b.price,
	   LineTotal = b.LineTotal,
	   VatSum = b.VatSum,
	   VatGroup = b.VatGroup
from OINV a 
	 inner join INV1 b ON a.DocEntry = b.DocEntry
where (a.BPLId = @BPLId)
  and a.DocDate between @DocDateFr and @DocDateTo
  
  Union all
  
  Select DocEntry  = b.BaseEntry,
		   LineNum = b.BaseLine,
	   CardCode  = a.CardCode,
       GrosProfSy = a.GrosProfSy * -1,
		 VatSumSy = a.VatSumSy * -1,
       DocTotalSy = a.DocTotalSy * -1,
       ItemCode   = b.ItemCode,
       Quantity = b.Quantity * -1,
       price = b.price * -1,
	   LineTotal = b.LineTotal * -1,
	   VatSum = b.VatSum * -1,
	   VatGroup = b.VatGroup
from ORIN a 
	 inner join RIN1 b ON a.DocEntry = b.DocEntry
where (a.BPLId = @BPLId)
  and a.DocDate between @DocDateFr and @DocDateTo
  AND b.BaseType = '13'
 ) t1
 group by t1.DocEntry,
		  t1.LineNum,
		  t1.CardCode,
		  t1.ItemCode,
		  t1.VatGroup
Having sum(t1.LineTotal) <> 0
) a
 inner join OITM c on a.ItemCode = c.ItemCode
 left  join [@PSH_MARK] f on c.U_Mark = f.Code
 left  join OCRD i on a.CardCode = i.CardCode	
 
  
 Union all
 
 select Cnt = '',
	   i.VatRegNum,
       i.E_Mail,
       taxgubun = (CASE WHEN b.VatGroup = 'A3' then 'P' /* 로칼=영세 */
                   ELSE 'T'   /* 내수=과세 */
                   END),
       seqn     = 'AC' + convert(nvarchar(10),a.DocEntry),
       ymd      = Convert(char(8),a.DocDate,112),
       amt      = a.GrosProfSy * -1,
	   vat      = a.VatSumSy * -1,
       tot      = a.DocTotalSy * -1,
       mmdd     = right(Convert(Char(8),a.DocDate,112) ,4),
       ItemCode   = b.ItemCode,
       ItemName   = c.FrgnName + (CASE WHEN c.ItmsGrpCod = '101' THEN ' [상품]'
                              ELSE ''
                              END),
       spec     = isnull(Convert(Nvarchar(100),c.U_Size),'') + ' ' +isnull(Convert(Nvarchar(10),f.Name),''),
       qty      = b.Quantity * -1,
       price    = b.price * -1,
	   damt     = b.LineTotal * -1,
	   dvat     = b.VatSum * -1,
       dbigo    = round((b.Quantity * c.U_UnWeight) / 1000,3) * -1,
       i.CardName,
       DocEntry = 'AC' + convert(nvarchar(10),a.DocEntry),
       LineNum = convert(nvarchar(10),b.LineNum)
from ORIN a 
	 inner join RIN1 b ON a.DocEntry = b.DocEntry
	 inner join OITM c on b.ItemCode = c.ItemCode
	 left  join [@PSH_MARK] f on c.U_Mark = f.Code
	 left  join OCRD i on a.CardCode = i.CardCode	
where (a.BPLId = @BPLId)
  and a.DocDate between @DocDateFr and @DocDateTo
  AND b.BaseType <> '13'
  


DECLARE CUR_1 CURSOR FOR
	Select DocEntry, LineNum
	  From #Temp01
	Order by DocEntry, LineNum
			
OPEN CUR_1
FETCH NEXT FROM CUR_1 INTO 	@DocEntry, @LineNum
WHILE	@@FETCH_STATUS = 0
BEGIN				    

	If @BeforeDoc <> @DocEntry Begin
		Update #Temp01 Set Cnt = Right('$$$$$' + Convert(Nvarchar(5),@cnt),5)
		Where DocEntry = @DocEntry
		
		Set @Cnt = @Cnt + 1  
	End
	
	
	If @BeforeDoc = @DocEntry Begin
		Update #Temp01
		   set cnt = '', VatRegNum = '', Email = '', taxgubun = '', Seqn = '', ymd = '', amt = 0, vat = 0, tot = 0
			   
		 where DocEntry = @DocEntry
		   and LineNum = @LineNum
		   
		
	End
	
	
	
	--If @BeforeDoc <> @OrdNum Begin
	--	Set @BeforeGONGSENG = 0
	--	Set @JEGONG = 0
	--End
	
	
	
	--Insert #Temp01 values (	@OrdNum, @ItemCode, @ItemName, @Selwt, @JISURANG, @CpCode, @CPName, @PQty, @OutQty, 
	--							@NQty, @GONGSENG, @DocDate, @JEGONG, 0, @SEOULDE, @SEOULJE, @Sequence)	
	Set @BeforeDoc = @DocEntry
	
	
FETCH NEXT FROM CUR_1 INTO 	@DocEntry, @LineNum
END	

CLOSE	CUR_1
DEALLOCATE CUR_1

 select cnt As '1구분자', 
	    VatRegNum As '2사업자번호',
		CardName As '상호',
	    '' As '3종사업장번호',
		Email As '4거래처이메일',
		'' As '5담당자' ,
		'' As '6담당자이메일',
		taxgubun As '7종류',
		Seqn As '8관리번호' ,
		ymd As '9일자',
		amt As '10공급가액',
		vat As '11세액',
		tot As '12합계',
		'' As '13비고1',
		'' As '14비고2',
		'' As '15비고3',
		'18' As '16청구영수', --청구
        '' As '17현금',
        '' As '18수표',
        '' As '19어음',
        '' As '20외상미수금',
        '' As '21기타번호1',
        '' As '22기타번호2',
        '' As '23기타번호3',
        '' As '24기타번호4',
        '' As '25수정코드',
		mmdd As '26월일',
		ItemCode As '27코드' ,
		ItemName As '28품목명',
		size As '29규격',
		Quantity As '30수량',
		Price As '31단가',
		damt As '32공급가',
		dvat As '33세액',
		bigo As '34비고',
		DocEntry, --송장번호
		LineNum  --송장순번	
  From	#Temp01
 Order by DocEntry, LineNum

End

--  EXEC [dbo].[PS_SD_SMARTBILL] '4','20110101','20110131'
